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

# =========================
# INPUT / OUTPUT LAYOUT
# =========================
# A = Action
# B = status_message
# C = error_message
# D = db_name
# E = db_id
# F = export_id
# G = export_name

INPUT_RANGE = "A:G"
OUTPUT_START_CELL = "D1"

FIXED_HEADERS = [
    "db_name",
    "db_id",
    "export_id",
    "export_name",
    "cron",
    "file_name",
    "export_protocol",
    "username",
    "cron_timezone",
    "destination",
    "host",
    "password",
    "raw_export_json",
    "field_json",
    "export_selector",
    "threshold"
]


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


def pad_rows(values, total_cols):
    padded = []
    for row in values:
        row = row + [""] * (total_cols - len(row))
        padded.append(row[:total_cols])
    return padded


# =========================
# SHEETS READ / WRITE
# =========================
def read_sheet_rows():
    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{INPUT_RANGE}"
    ).execute()

    values = result.get("values", [])
    if not values:
        return []

    values = pad_rows(values, 7)
    data_rows = values[1:] if len(values) > 1 else []

    rows = []
    for idx, row in enumerate(data_rows, start=2):
        action = str(row[0]).strip()
        status_message = str(row[1]).strip()
        error_message = str(row[2]).strip()
        db_name = str(row[3]).strip()
        db_id = str(row[4]).strip()
        export_id = str(row[5]).strip()
        export_name = str(row[6]).strip()

        if (
            action == "" and
            status_message == "" and
            error_message == "" and
            db_name == "" and
            db_id == "" and
            export_id == "" and
            export_name == ""
        ):
            continue

        rows.append({
            "sheet_row": idx,
            "action": action,
            "status_message": status_message,
            "error_message": error_message,
            "db_name": db_name,
            "db_id": db_id,
            "export_id": export_id,
            "export_name": export_name
        })

    return rows


def clear_output():
    service = get_sheets_service()
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!D:ZZZ"
    ).execute()


def write_output_table(headers, rows):
    service = get_sheets_service()

    values = [headers] + rows

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{OUTPUT_START_CELL}",
        valueInputOption="RAW",
        body={"values": values}
    ).execute()


def update_status_columns(status_rows):
    service = get_sheets_service()

    if not status_rows:
        return

    status_rows_sorted = sorted(status_rows, key=lambda x: x["sheet_row"])
    start_row = status_rows_sorted[0]["sheet_row"]
    end_row = status_rows_sorted[-1]["sheet_row"]

    row_map = {item["sheet_row"]: item for item in status_rows_sorted}

    values_b = []
    values_c = []

    for row_num in range(start_row, end_row + 1):
        item = row_map.get(row_num, {})
        values_b.append([item.get("status_message", "")])
        values_c.append([item.get("error_message", "")])

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!B{start_row}:B{end_row}",
        valueInputOption="RAW",
        body={"values": values_b}
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!C{start_row}:C{end_row}",
        valueInputOption="RAW",
        body={"values": values_c}
    ).execute()


# =========================
# API
# =========================
def get_exports(db_id):
    url = f"{service_path}/dbs/{db_id}/exports"

    resp = requests.get(
        url,
        headers=headers,
        verify=False,
        timeout=60
    )

    payload = safe_json(resp)

    if resp.status_code != 200:
        raise Exception(f"HTTP {resp.status_code}: {dumps_safe(payload)}")

    if isinstance(payload, dict) and payload.get("status") == "fail":
        raise Exception(f"Feedonomics fail payload: {dumps_safe(payload)}")

    if not isinstance(payload, list):
        raise Exception(f"Unexpected payload type: {type(payload)} | {payload}")

    return payload


def get_schedule_optional(db_id, export_id):
    try:
        url = f"{service_path}/dbs/{db_id}/exports/{export_id}/schedule"
        resp = requests.get(
            url,
            headers=headers,
            verify=False,
            timeout=60
        )

        if resp.status_code == 200:
            payload = safe_json(resp)
            if isinstance(payload, dict):
                return payload
    except Exception:
        pass

    return {}


def find_export(exports, export_id=None, export_name=None):
    if export_id:
        export_id_str = str(export_id).strip()
        for exp in exports:
            if str(exp.get("id", "")).strip() == export_id_str:
                return exp

    if export_name:
        export_name_lower = str(export_name).strip().lower()
        for exp in exports:
            if str(exp.get("name", "")).strip().lower() == export_name_lower:
                return exp

    return None


# =========================
# EXTRACTION
# =========================
def extract_cron(exp, sch):
    schedule_block = exp.get("schedule", {})
    if not isinstance(schedule_block, dict):
        schedule_block = {}

    return first(
        sch.get("cron") if isinstance(sch, dict) else "",
        exp.get("cron"),
        exp.get("schedule_cron"),
        schedule_block.get("cron")
    )


def extract_cron_timezone(exp, sch):
    schedule_block = exp.get("schedule", {})
    if not isinstance(schedule_block, dict):
        schedule_block = {}

    return first(
        sch.get("cron_timezone") if isinstance(sch, dict) else "",
        sch.get("timezone") if isinstance(sch, dict) else "",
        exp.get("cron_timezone"),
        exp.get("timezone"),
        exp.get("time_zone"),
        schedule_block.get("cron_timezone"),
        schedule_block.get("timezone")
    )


def extract_fields_json(exp):
    export_fields = exp.get("export_fields", [])
    if not isinstance(export_fields, list):
        export_fields = []
    return dumps_safe(export_fields)


def extract_export_selector(exp):
    return first(
        exp.get("export_selector"),
        exp.get("selector"),
        exp.get("select_rule"),
        exp.get("filter"),
        exp.get("where")
    )


def extract_threshold(exp):
    return first(
        exp.get("threshold"),
        exp.get("load_threshold"),
        exp.get("min_threshold"),
        exp.get("max_threshold")
    )


def extract_field_triplets(exp):
    export_fields = exp.get("export_fields", [])
    if not isinstance(export_fields, list):
        return []

    triplets = []
    for f in export_fields:
        if not isinstance(f, dict):
            continue

        triplets.append([
            str(first(f.get("id"), f.get("export_field_id"), f.get("field_id"))).strip(),
            str(f.get("field_name", "")).strip(),
            str(f.get("export_field_name", "")).strip()
        ])

    return triplets


def build_dynamic_headers(max_fields):
    headers = FIXED_HEADERS[:]
    for i in range(1, max_fields + 1):
        headers.extend([
            f"id_{i}",
            f"field_name_{i}",
            f"export_field_name_{i}"
        ])
    return headers


def build_output_row(input_row, exp=None, sch=None, max_fields=0):
    base_row = [
        input_row["db_name"],
        input_row["db_id"],
        input_row["export_id"],
        input_row["export_name"],
        extract_cron(exp, sch),
        first(exp.get("file_name"), exp.get("name")),
        first(exp.get("protocol"), exp.get("type"), exp.get("method")),
        first(exp.get("username"), exp.get("ftp_user"), exp.get("user")),
        extract_cron_timezone(exp, sch),
        first(exp.get("destination"), exp.get("ftp_path"), exp.get("path")),
        first(exp.get("host"), exp.get("ftp_host"), exp.get("server")),
        first(exp.get("password"), exp.get("ftp_password")),
        dumps_safe(exp),
        extract_fields_json(exp),
        extract_export_selector(exp),
        extract_threshold(exp)
    ]

    triplets = extract_field_triplets(exp)

    flat_triplets = []
    for triplet in triplets:
        flat_triplets.extend(triplet)

    expected_triplet_cells = max_fields * 3
    flat_triplets += [""] * (expected_triplet_cells - len(flat_triplets))

    return base_row + flat_triplets


def build_error_row(input_row, max_fields=0):
    base_row = [
        input_row["db_name"],
        input_row["db_id"],
        input_row["export_id"],
        input_row["export_name"],
        "", "", "", "", "", "", "", "", "", "", "", ""
    ]
    return base_row + ([""] * (max_fields * 3))


# =========================
# MAIN
# =========================
def main():
    print("📥 Leyendo filas de Update Exports...")
    input_rows = read_sheet_rows()

    if not input_rows:
        print("⚠️ No se encontraron filas válidas.")
        return

    print(f"🔎 Filas válidas encontradas: {len(input_rows)}")

    prepared_rows = []
    status_updates = []
    exports_cache = {}
    max_fields = 0

    for row in input_rows:
        db_name = row["db_name"]
        db_id_raw = row["db_id"]
        export_id_raw = row["export_id"]
        export_name = row["export_name"]
        sheet_row = row["sheet_row"]

        try:
            if not db_id_raw or not export_id_raw:
                raise Exception("db_id y export_id son requeridos")

            db_id = clean_int(db_id_raw)
            export_id = clean_int(export_id_raw)

            if db_id in exports_cache:
                exports = exports_cache[db_id]
            else:
                print(f"🚀 Consultando exports | DB {db_id}")
                exports = get_exports(db_id)
                exports_cache[db_id] = exports

            exp = find_export(exports, export_id=export_id, export_name=export_name)

            if not exp:
                raise Exception("Export not found")

            sch = get_schedule_optional(db_id, export_id)
            field_count = len(extract_field_triplets(exp))
            max_fields = max(max_fields, field_count)

            prepared_rows.append({
                "type": "success",
                "input_row": row,
                "exp": exp,
                "sch": sch
            })

            status_updates.append({
                "sheet_row": sheet_row,
                "status_message": f"SUCCESS | {field_count} field(s)",
                "error_message": ""
            })

            print(f"✅ OK | {db_name} | export_id {export_id} | fields={field_count}")

        except Exception as e:
            prepared_rows.append({
                "type": "error",
                "input_row": row
            })

            status_updates.append({
                "sheet_row": sheet_row,
                "status_message": "ERROR",
                "error_message": str(e)
            })

            print(f"❌ ERROR | {db_name} | {str(e)}")

    headers = build_dynamic_headers(max_fields)

    output_rows = []
    for item in prepared_rows:
        if item["type"] == "success":
            output_rows.append(
                build_output_row(
                    item["input_row"],
                    exp=item["exp"],
                    sch=item["sch"],
                    max_fields=max_fields
                )
            )
        else:
            output_rows.append(
                build_error_row(
                    item["input_row"],
                    max_fields=max_fields
                )
            )

    print("🧹 Limpiando salida anterior desde columna D...")
    clear_output()

    print("✍️ Escribiendo resultados por filas desde D...")
    write_output_table(headers, output_rows)

    print("✍️ Actualizando status_message y error_message...")
    update_status_columns(status_updates)

    print("✅ Proceso completado")
    print(f"   - Filas procesadas: {len(input_rows)}")
    print(f"   - Filas escritas: {len(output_rows)}")
    print(f"   - Máximo de fields encontrados en un export: {max_fields}")


if __name__ == "__main__":
    main()
