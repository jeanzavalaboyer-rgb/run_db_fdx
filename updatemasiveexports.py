import os
import requests
import urllib3
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# CONFIG
# =========================
api_key = os.environ["FEEDONOMICS_API_KEY"]
service_path = "https://meta.feedonomics.com/api.php"
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

headers = {
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Content-Type": "application/json"
}

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw")
SHEET_NAME = os.getenv("SHEET_NAME", "Update Exports")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

BLOCK_WIDTH = 4
BLOCK_GAP = 1


# =========================
# GOOGLE SHEETS
# =========================
def get_sheets_service():
    service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def col_to_letter(col_num: int) -> str:
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def read_sheet_matrix():
    service = get_sheets_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A1:ZZ500"
    ).execute()
    return result.get("values", [])


def get_cell(matrix, row_idx_1based, col_idx_1based):
    r = row_idx_1based - 1
    c = col_idx_1based - 1

    if r < 0 or r >= len(matrix):
        return ""

    row = matrix[r]
    if c < 0 or c >= len(row):
        return ""

    return str(row[c]).strip()


def is_range_empty(matrix, row_from, row_to, col_from, col_to):
    for r in range(row_from, row_to + 1):
        for c in range(col_from, col_to + 1):
            if get_cell(matrix, r, c) != "":
                return False
    return True


def read_input_blocks(matrix):
    """
    Lee bloques:
    A:D, F:I, K:N, P:S, ...
    Usa fila 1 como header y fila 2 como input del usuario.
    """
    blocks = []
    col = 1

    while True:
        current_block_empty = is_range_empty(matrix, 1, 2, col, col + 3)
        next_two_cols_empty = is_range_empty(matrix, 1, 2, col + 4, col + 5)

        if current_block_empty and next_two_cols_empty:
            break

        if not current_block_empty:
            blocks.append({
                "start_col": col,
                "db_name": get_cell(matrix, 2, col),
                "db_id": get_cell(matrix, 2, col + 1),
                "export_id": get_cell(matrix, 2, col + 2),
                "export_name": get_cell(matrix, 2, col + 3),
            })

        col += BLOCK_WIDTH + BLOCK_GAP

    return blocks


# =========================
# HELPERS
# =========================
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


def normalize_action(value):
    return str(value).strip().upper()


def ensure_json_field(payload, field_name, default):
    value = payload.get(field_name)

    if value is None or value == "":
        payload[field_name] = default
        return

    if isinstance(value, str):
        try:
            payload[field_name] = json.loads(value)
        except Exception:
            payload[field_name] = default


def values_differ(a, b):
    return str(first(a, "")).strip() != str(first(b, "")).strip()


# =========================
# API
# =========================
def update_export(db_id, export_id, payload):
    url = f"{service_path}/dbs/{db_id}/exports/{export_id}"

    resp = requests.put(
        url,
        headers=headers,
        json=payload,
        verify=False,
        timeout=120
    )

    payload_resp = safe_json(resp)

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}: {dumps_safe(payload_resp)}"

    return True, dumps_safe(payload_resp)


def update_export_schedule(db_id, export_id, payload):
    url = f"{service_path}/dbs/{db_id}/exports/{export_id}/schedule"

    resp = requests.put(
        url,
        headers=headers,
        json=payload,
        verify=False,
        timeout=120
    )

    payload_resp = safe_json(resp)

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}: {dumps_safe(payload_resp)}"

    return True, dumps_safe(payload_resp)


# =========================
# SHEET PARSING
# =========================
def parse_marked_field_updates_and_raw_json(matrix, block):
    start_col = block["start_col"]
    max_scan_row = 500

    raw_export_json = ""
    updates = []

    row = 3
    while row <= max_scan_row:
        c1 = get_cell(matrix, row, start_col)
        c2 = get_cell(matrix, row, start_col + 1)
        c3 = get_cell(matrix, row, start_col + 2)
        c4 = get_cell(matrix, row, start_col + 3)

        if c1.lower() == "raw_export_json":
            raw_export_json = c2

        if (
            c1.lower() == "id"
            and c3.lower() == "action"
            and get_cell(matrix, row + 1, start_col).lower() == "field_name"
            and get_cell(matrix, row + 1, start_col + 2).lower() == "export_field_name"
        ):
            action = normalize_action(c4)

            if action == "Y":
                updates.append({
                    "id": c2,
                    "field_name": get_cell(matrix, row + 1, start_col + 1),
                    "export_field_name": get_cell(matrix, row + 1, start_col + 3),
                    "action_row": row
                })

            row += 2
            continue

        row += 1

    return raw_export_json, updates


def parse_editable_overrides(matrix, block):
    start_col = block["start_col"]
    max_scan_row = 80

    editable = {
        "cron": "",
        "cron_timezone": "",
        "file_name": "",
        "destination": "",
        "export_protocol": "",
        "host": "",
        "username": "",
        "password": ""
    }

    row = 3
    while row <= max_scan_row:
        left_label = get_cell(matrix, row, start_col).strip().lower()
        left_value = get_cell(matrix, row, start_col + 1)
        right_label = get_cell(matrix, row, start_col + 2).strip().lower()
        right_value = get_cell(matrix, row, start_col + 3)

        if left_label in editable:
            editable[left_label] = left_value

        if right_label in editable:
            editable[right_label] = right_value

        row += 1

    return editable


# =========================
# PAYLOAD BUILD
# =========================
def build_export_update_payload_from_raw(raw_export_json, marked_updates):
    if not raw_export_json.strip():
        raise Exception("raw_export_json vacío")

    try:
        payload = json.loads(raw_export_json)
    except Exception as e:
        raise Exception(f"No se pudo parsear raw_export_json: {str(e)}")

    if not isinstance(payload, dict):
        raise Exception("raw_export_json no es un objeto JSON válido")

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

    export_fields = payload.get("export_fields", [])
    if not isinstance(export_fields, list):
        raise Exception("raw_export_json.export_fields no viene como lista")

    updates_by_id = {
        str(item["id"]).strip(): item
        for item in marked_updates
        if str(item.get("id", "")).strip()
    }

    export_fields_payload = {}
    field_counter = 1

    for field in export_fields:
        if not isinstance(field, dict):
            continue

        field_copy = dict(field)

        field_id = str(first(
            field_copy.get("id"),
            field_copy.get("export_field_id"),
            field_copy.get("field_id")
        )).strip()

        if field_id in updates_by_id:
            marked = updates_by_id[field_id]
            field_copy["field_name"] = marked.get("field_name", field_copy.get("field_name", ""))
            field_copy["export_field_name"] = marked.get("export_field_name", field_copy.get("export_field_name", ""))

        export_fields_payload[f"field{field_counter}"] = field_copy
        field_counter += 1

    payload["export_fields"] = export_fields_payload
    return payload


def apply_sheet_overrides_to_payload(payload, overrides):
    if overrides.get("file_name", "").strip() != "":
        payload["file_name"] = overrides["file_name"]

    if overrides.get("destination", "").strip() != "":
        payload["destination"] = overrides["destination"]

    if overrides.get("export_protocol", "").strip() != "":
        payload["protocol"] = overrides["export_protocol"]

    if overrides.get("host", "").strip() != "":
        payload["host"] = overrides["host"]

    if overrides.get("username", "").strip() != "":
        payload["username"] = overrides["username"]

    if overrides.get("password", "").strip() != "":
        payload["password"] = overrides["password"]

    cron_value = str(first(overrides.get("cron"), "")).strip()
    cron_timezone_value = str(first(overrides.get("cron_timezone"), "")).strip()

    # Mantengo la lógica EXACTA del original
    if cron_value == "":
        payload["cron"] = "null"
        payload["cron_timezone"] = ""
        payload["paused"] = 1
    else:
        payload["cron"] = cron_value
        if cron_timezone_value != "":
            payload["cron_timezone"] = cron_timezone_value

    return payload


def build_schedule_payload_from_cron(cron_value, cron_timezone_value="", current_paused="0"):
    cron_str = str(first(cron_value, "")).strip()
    cron_timezone_str = str(first(cron_timezone_value, "")).strip()

    # Si cron está vacío, mandar null explícito
    if not cron_str:
        return {
            "cron": None,
            "cron_timezone": None,
            "paused": 1
        }

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
        "paused": int(str(first(current_paused, "0")).strip() or "0")
    }

    if cron_timezone_str != "":
        payload["cron_timezone"] = cron_timezone_str

    return payload


# =========================
# WRITE STATUS BACK TO SHEET
# =========================
def write_status_cell(service, row_1based, col_1based, value):
    col_letter = col_to_letter(col_1based)
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{col_letter}{row_1based}",
        valueInputOption="RAW",
        body={"values": [[value]]}
    ).execute()


# =========================
# MAIN
# =========================
def main():
    service = get_sheets_service()
    matrix = read_sheet_matrix()
    blocks = read_input_blocks(matrix)

    if not blocks:
        print("⚠️ No se encontraron bloques de entrada.")
        return

    for block in blocks:
        db_name = block["db_name"]
        db_id_raw = block["db_id"]
        export_id_raw = block["export_id"]
        start_col = block["start_col"]

        if not db_name and not db_id_raw and not export_id_raw and not block["export_name"]:
            continue

        try:
            if not db_id_raw or not export_id_raw:
                print(f"⏭️ SKIPPED | {db_name} | faltan db_id/export_id")
                continue

            db_id = clean_int(db_id_raw)
            export_id = clean_int(export_id_raw)

            raw_export_json, marked_updates = parse_marked_field_updates_and_raw_json(matrix, block)
            overrides = parse_editable_overrides(matrix, block)

            raw_payload = json.loads(raw_export_json) if raw_export_json.strip() else {}
            if not isinstance(raw_payload, dict):
                raise Exception("raw_export_json no es un objeto JSON válido")

            current_cron = first(raw_payload.get("cron"), "")
            current_paused = first(raw_payload.get("paused"), "0")

            cron_override = str(first(overrides.get("cron"), "")).strip()
            cron_changed = values_differ(cron_override, current_cron)

            non_schedule_override_present = any([
                overrides.get("file_name", "").strip(),
                overrides.get("destination", "").strip(),
                overrides.get("export_protocol", "").strip(),
                overrides.get("host", "").strip(),
                overrides.get("username", "").strip(),
                overrides.get("password", "").strip(),
            ])

            if not marked_updates and not cron_changed and not non_schedule_override_present:
                print(f"⏭️ SKIPPED | {db_name} | export_id {export_id} | no hay cambios")
                continue

            export_ok = True
            export_response_text = "SKIPPED"

            if marked_updates or non_schedule_override_present or cron_changed:
                payload = build_export_update_payload_from_raw(raw_export_json, marked_updates)
                payload = apply_sheet_overrides_to_payload(payload, overrides)

                print("=== EXPORT PAYLOAD ENVIADO ===")
                print(json.dumps(payload, indent=2, ensure_ascii=False)[:20000])

                export_ok, export_response_text = update_export(db_id, export_id, payload)

            schedule_ok = True
            schedule_response_text = "SKIPPED"

            if cron_changed:
                schedule_payload = build_schedule_payload_from_cron(
                    cron_value=overrides.get("cron"),
                    cron_timezone_value=overrides.get("cron_timezone"),
                    current_paused=current_paused
                )

                print("=== SCHEDULE PAYLOAD ENVIADO ===")
                print(json.dumps(schedule_payload, indent=2, ensure_ascii=False))

                schedule_ok, schedule_response_text = update_export_schedule(
                    db_id=db_id,
                    export_id=export_id,
                    payload=schedule_payload
                )

            final_ok = export_ok and schedule_ok
            combined_response = (
                f"EXPORT: {export_response_text}\n\n"
                f"SCHEDULE: {schedule_response_text}"
            )

            for item in marked_updates:
                action_row = item["action_row"]
                write_status_cell(service, action_row, start_col + 4, "SUCCESS" if final_ok else "ERROR")
                write_status_cell(service, action_row + 1, start_col + 4, combined_response[:50000])

            if not marked_updates and (cron_changed or non_schedule_override_present):
                write_status_cell(service, 4, start_col + 4, "SUCCESS" if final_ok else "ERROR")
                write_status_cell(service, 5, start_col + 4, combined_response[:50000])

            if final_ok:
                print(f"✅ UPDATED | {db_name} | export_id {export_id}")
            else:
                print(f"❌ ERROR | {db_name} | export_id {export_id} | {combined_response}")

        except Exception as e:
            print(f"❌ ERROR | {db_name} | export_id {export_id_raw} | {str(e)}")


if __name__ == "__main__":
    main()
