import os
import requests
import urllib3
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =====================================================
# CONFIG
# =====================================================
api_key = os.environ["FEEDONOMICS_API_KEY"]
service_path = "https://meta.feedonomics.com/api.php"
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

headers = {
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Content-Type": "application/json"
}

SPREADSHEET_ID = os.getenv(
    "SPREADSHEET_ID",
    "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw"
)

SHEET_NAME = os.getenv(
    "SHEET_NAME",
    "Update Exports"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]

# =====================================================
# SHEET LAYOUT
# =====================================================
# A = Action
# B = status_message
# C = error_message
# D = db_name
# E = db_id
# F = export_id
# G = export_name
# H = cron
# I = file_name
# J = export_protocol
# K = username
# L = cron_timezone
# M = destination
# N = host
# O = password
# P = raw_export_json
# Q = field_json
# R = export_selector
# S = threshold
# T,U,V = id_1, field_name_1, export_field_name_1
# W,X,Y = id_2, field_name_2, export_field_name_2
# ...

INPUT_RANGE = "A:ZZ"

# =====================================================
# GOOGLE SHEETS
# =====================================================
def get_sheets_service():
    service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)

    creds = Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


# =====================================================
# HELPERS
# =====================================================
def safe_json(resp):
    try:
        return resp.json()
    except Exception:
        return {"_raw": resp.text}


def first(*vals):
    for v in vals:
        if v is not None and str(v).strip() != "":
            return v
    return ""


def clean_int(v):
    return int(float(str(v).strip()))


def dumps_safe(obj):
    try:
        return json.dumps(obj, ensure_ascii=False)
    except Exception:
        return str(obj)


def normalize_action(v):
    return str(v).strip().upper()


def col_to_letter(col_num):
    result = ""
    while col_num > 0:
        col_num, rem = divmod(col_num - 1, 26)
        result = chr(65 + rem) + result
    return result


def values_differ(a, b):
    return str(first(a, "")).strip() != str(first(b, "")).strip()


def ensure_json_field(payload, field_name, default):
    value = payload.get(field_name)

    if value is None or value == "":
        payload[field_name] = default
        return

    if isinstance(value, str):
        try:
            parsed = json.loads(value)
            payload[field_name] = parsed
        except Exception:
            payload[field_name] = default


# =====================================================
# SHEETS
# =====================================================
def read_sheet_matrix():
    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{INPUT_RANGE}"
    ).execute()

    return result.get("values", [])


def get_cell(matrix, row_1, col_1):
    r = row_1 - 1
    c = col_1 - 1

    if r < 0 or r >= len(matrix):
        return ""

    row = matrix[r]

    if c < 0 or c >= len(row):
        return ""

    return str(row[c]).strip()


def write_cell(row_1, col_1, value):
    service = get_sheets_service()

    col = col_to_letter(col_1)

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{col}{row_1}",
        valueInputOption="RAW",
        body={"values": [[value]]}
    ).execute()


# =====================================================
# API
# =====================================================
def update_export(db_id, export_id, payload):
    url = f"{service_path}/dbs/{db_id}/exports/{export_id}"

    resp = requests.put(
        url,
        headers=headers,
        json=payload,
        verify=False,
        timeout=120
    )

    data = safe_json(resp)

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}: {dumps_safe(data)}"

    return True, dumps_safe(data)


def update_export_schedule(db_id, export_id, payload):
    url = f"{service_path}/dbs/{db_id}/exports/{export_id}/schedule"

    resp = requests.put(
        url,
        headers=headers,
        json=payload,
        verify=False,
        timeout=120
    )

    data = safe_json(resp)

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}: {dumps_safe(data)}"

    return True, dumps_safe(data)


def delete_export(db_id, export_id):
    url = f"{service_path}/dbs/{db_id}/exports/{export_id}"

    resp = requests.delete(
        url,
        headers=headers,
        verify=False,
        timeout=120
    )

    data = safe_json(resp)

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}: {dumps_safe(data)}"

    return True, dumps_safe(data)


# =====================================================
# FIELD GROUPS
# =====================================================
def parse_horizontal_fields(row_values):
    """
    Desde columna T = 20 en adelante:
    grupos de 3 columnas:
    id / field_name / export_field_name

    Regla:
    - si hay un bloque vacío en el medio, NO se corta la lectura
    - solo se ignora ese bloque
    - así no se pierden los campos posteriores
    """
    groups = []
    col = 20
    empty_streak = 0
    max_col = 500

    while col <= max_col:
        field_id = str(row_values.get(col, "")).strip()
        field_name = str(row_values.get(col + 1, "")).strip()
        export_field_name = str(row_values.get(col + 2, "")).strip()

        is_empty_group = (
            field_id == "" and
            field_name == "" and
            export_field_name == ""
        )

        if is_empty_group:
            empty_streak += 1
        else:
            empty_streak = 0
            groups.append({
                "id": field_id,
                "field_name": field_name,
                "export_field_name": export_field_name
            })

        # corta solo si ya hay varios bloques vacíos seguidos al final
        if empty_streak >= 5:
            break

        col += 3

    return groups


# =====================================================
# BUILD PAYLOAD
# =====================================================
def build_payload_from_raw(raw_json):
    if not raw_json.strip():
        raise Exception("raw_export_json vacío")

    payload = json.loads(raw_json)

    if not isinstance(payload, dict):
        raise Exception("raw_export_json inválido")

    for key in [
        "created_at", "updated_at", "last_run_at", "next_run_time",
        "timestamp", "worker_pid", "worker_hostname", "summary_worker_hostname",
        "summary_worker_pid", "time_running", "average_chunk_memory",
        "average_chunk_duration", "last_run_duration", "last_run_status",
        "last_run_total_rows", "last_run_time", "paused_at", "worker_stage",
        "health_status", "general_export_file_url"
    ]:
        payload.pop(key, None)

    ensure_json_field(payload, "strip_characters", ["\r", "\n", "\t"])
    ensure_json_field(payload, "protocol_info", {})
    ensure_json_field(payload, "sortable_fields", {})
    ensure_json_field(payload, "deduplicate_field_name", [])
    ensure_json_field(payload, "tags", [])

    return payload


def apply_overrides(payload, row):
    if row["file_name"]:
        payload["file_name"] = row["file_name"]

    if row["destination"]:
        payload["destination"] = row["destination"]

    if row["protocol"]:
        payload["protocol"] = row["protocol"]

    if row["host"]:
        payload["host"] = row["host"]

    if row["username"]:
        payload["username"] = row["username"]

    if row["password"]:
        payload["password"] = row["password"]

    if row["export_selector"]:
        payload["export_selector"] = row["export_selector"]

    if row["threshold"]:
        payload["threshold"] = row["threshold"]

    cron_value = str(first(row["cron"], "")).strip()
    cron_timezone_value = str(first(row["cron_timezone"], "")).strip()

    if cron_value == "":
        payload["cron"] = None
        payload["cron_timezone"] = None if cron_timezone_value == "" else cron_timezone_value
        payload["paused"] = 1
    else:
        payload["cron"] = cron_value
        if cron_timezone_value != "":
            payload["cron_timezone"] = cron_timezone_value
        payload["paused"] = 0

    return payload


def replace_export_fields(payload, groups):
    original = payload.get("export_fields", [])

    if isinstance(original, dict):
        original = list(original.values())

    if not isinstance(original, list):
        original = []

    existing_by_id = {}

    for item in original:
        fid = str(first(
            item.get("id"),
            item.get("field_id"),
            item.get("export_field_id")
        )).strip()

        if fid:
            existing_by_id[fid] = dict(item)

    final = {}
    idx = 1

    for g in groups:
        fid = str(g.get("id", "")).strip()
        field_name = str(g.get("field_name", "")).strip()
        export_field_name = str(g.get("export_field_name", "")).strip()

        # ignora bloques totalmente vacíos
        if fid == "" and field_name == "" and export_field_name == "":
            continue

        # para mantener o actualizar un field existente, debe tener id
        if fid == "":
            continue

        item = existing_by_id.get(fid, {"id": fid})
        item["id"] = fid
        item["field_name"] = field_name
        item["export_field_name"] = export_field_name

        final[f"field{idx}"] = item
        idx += 1

    payload["export_fields"] = final
    return payload


# =====================================================
# SCHEDULE
# =====================================================
def build_schedule_payload_from_cron(cron_value, cron_timezone_value="", current_paused="0"):
    cron_str = str(first(cron_value, "")).strip()
    cron_timezone_str = str(first(cron_timezone_value, "")).strip()

    if not cron_str:
        payload = {
            "cron": None,
            "paused": 1
        }
        if cron_timezone_str != "":
            payload["cron_timezone"] = cron_timezone_str
        else:
            payload["cron_timezone"] = None
        return payload

    parts = cron_str.split()
    if len(parts) != 5:
        raise Exception(f"Cron inválido: '{cron_str}'. Debe tener 5 partes.")

    minute, hour, day_of_month, month, weekday = parts

    if day_of_month != "*" or month != "*":
        raise Exception(
            f"Cron no soportado para /schedule: '{cron_str}'. "
            f"Solo se soporta day_of_month='*' y month='*'."
        )

    weekday_map = {
        "SUN": "0",
        "MON": "1",
        "TUE": "2",
        "WED": "3",
        "THU": "4",
        "FRI": "5",
        "SAT": "6"
    }

    weekday_upper = weekday.upper()
    day_value = weekday_map.get(weekday_upper, weekday)

    payload = {
        "day": str(day_value),
        "hour": str(hour),
        "minute": str(minute),
        "paused": 0
    }

    if cron_timezone_str != "":
        payload["cron_timezone"] = cron_timezone_str

    return payload


# =====================================================
# MAIN
# =====================================================
def main():
    matrix = read_sheet_matrix()

    if len(matrix) <= 1:
        print("⚠️ No rows found")
        return

    for row_num in range(2, len(matrix) + 1):
        action = normalize_action(get_cell(matrix, row_num, 1))

        if action not in ["UPDATE", "DELETE"]:
            continue

        try:
            db_name = get_cell(matrix, row_num, 4)
            db_id_raw = get_cell(matrix, row_num, 5)
            export_id_raw = get_cell(matrix, row_num, 6)

            if not db_id_raw or not export_id_raw:
                raise Exception("Faltan db_id o export_id")

            db_id = clean_int(db_id_raw)
            export_id = clean_int(export_id_raw)

            if action == "DELETE":
                ok, msg = delete_export(db_id, export_id)

                write_cell(row_num, 2, "SUCCESS" if ok else "ERROR")
                write_cell(row_num, 3, msg[:50000])

                if ok:
                    print(f"🗑️ DELETED | row {row_num} | {db_name} | export_id {export_id}")
                else:
                    print(f"❌ DELETE ERROR | row {row_num} | {db_name} | {msg}")
                continue

            row = {
                "db_name": db_name,
                "db_id": db_id_raw,
                "export_id": export_id_raw,
                "export_name": get_cell(matrix, row_num, 7),
                "cron": get_cell(matrix, row_num, 8),
                "file_name": get_cell(matrix, row_num, 9),
                "protocol": get_cell(matrix, row_num, 10),
                "username": get_cell(matrix, row_num, 11),
                "cron_timezone": get_cell(matrix, row_num, 12),
                "destination": get_cell(matrix, row_num, 13),
                "host": get_cell(matrix, row_num, 14),
                "password": get_cell(matrix, row_num, 15),
                "raw_json": get_cell(matrix, row_num, 16),
                "field_json": get_cell(matrix, row_num, 17),
                "export_selector": get_cell(matrix, row_num, 18),
                "threshold": get_cell(matrix, row_num, 19),
            }

            raw_payload = json.loads(row["raw_json"]) if row["raw_json"].strip() else {}
            if not isinstance(raw_payload, dict):
                raise Exception("raw_export_json no es un objeto JSON válido")

            current_cron = first(raw_payload.get("cron"), "")
            current_paused = first(raw_payload.get("paused"), "0")

            cron_override = str(first(row["cron"], "")).strip()
            cron_changed = values_differ(cron_override, current_cron)

            row_values = {c: get_cell(matrix, row_num, c) for c in range(1, 501)}
            groups = parse_horizontal_fields(row_values)

            non_schedule_override_present = any([
                row["file_name"].strip(),
                row["destination"].strip(),
                row["protocol"].strip(),
                row["host"].strip(),
                row["username"].strip(),
                row["password"].strip(),
                row["export_selector"].strip(),
                row["threshold"].strip(),
                len(groups) > 0
            ])

            if not cron_changed and not non_schedule_override_present:
                write_cell(row_num, 2, "SKIPPED")
                write_cell(row_num, 3, "No hay cambios")
                print(f"⏭️ SKIPPED row {row_num} | no hay cambios")
                continue

            export_ok = True
            export_msg = "SKIPPED"

            if non_schedule_override_present or cron_changed:
                payload = build_payload_from_raw(row["raw_json"])
                payload = apply_overrides(payload, row)
                payload = replace_export_fields(payload, groups)

                print("=== EXPORT UPDATE ===")
                print(json.dumps(payload, indent=2, ensure_ascii=False)[:20000])

                export_ok, export_msg = update_export(
                    db_id,
                    export_id,
                    payload
                )

            schedule_ok = True
            schedule_msg = "SKIPPED"

            if cron_changed:
                schedule_payload = build_schedule_payload_from_cron(
                    cron_value=row["cron"],
                    cron_timezone_value=row["cron_timezone"],
                    current_paused=current_paused
                )

                print("=== SCHEDULE UPDATE ===")
                print(json.dumps(schedule_payload, indent=2, ensure_ascii=False))

                schedule_ok, schedule_msg = update_export_schedule(
                    db_id,
                    export_id,
                    schedule_payload
                )

            final_ok = export_ok and schedule_ok

            combined_msg = (
                f"EXPORT={export_msg[:4000]} | "
                f"SCHEDULE={schedule_msg[:4000]}"
            )

            write_cell(row_num, 2, "SUCCESS" if final_ok else "ERROR")
            write_cell(row_num, 3, combined_msg[:50000])

            if final_ok:
                print(f"✅ UPDATED row {row_num}")
            else:
                print(f"❌ row {row_num} | {combined_msg}")

        except Exception as e:
            write_cell(row_num, 2, "ERROR")
            write_cell(row_num, 3, str(e))
            print(f"❌ row {row_num} | {str(e)}")


if __name__ == "__main__":
    main()
