import os
import requests
import urllib3
import json
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# CONFIG
# =========================
api_key = os.environ["FEEDONOMICS_API_KEY"]
service_path = os.getenv("FEEDONOMICS_SERVICE_PATH", "https://meta.feedonomics.com/api.php")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

headers = {
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Content-Type": "application/json"
}

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw")
SHEET_NAME = os.getenv("SHEET_NAME", "Update Exports")
AVAILABLE_FIELDS_SHEET = os.getenv("AVAILABLE_FIELDS_SHEET", "Available Fields")

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


def ensure_sheet_exists(service, sheet_name):
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID
    ).execute()

    existing = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]
    if sheet_name in existing:
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet_name
                        }
                    }
                }
            ]
        }
    ).execute()


def read_sheet_matrix():
    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A1:ZZ400"
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

    Regla de parada:
    si el bloque actual está vacío y además las siguientes 2 columnas
    a la derecha están vacías, detiene el loop.
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


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


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


def get_db_fields(db_id: int):
    url = f"{service_path}/dbs/{db_id}/db_fields"
    resp = requests.get(url, headers=headers, verify=False, timeout=60)

    if resp.status_code != 200:
        return None, (resp.status_code, resp.text)

    payload = safe_json(resp)

    if isinstance(payload, dict) and payload.get("status") == "fail":
        return None, (resp.status_code, str(payload))

    if not isinstance(payload, list):
        return None, (resp.status_code, f"Unexpected payload: {type(payload)} | {payload}")

    return payload, None


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


def extract_field_rows(exp):
    export_fields = exp.get("export_fields", [])
    if not isinstance(export_fields, list):
        return []

    rows = []
    for f in export_fields:
        if not isinstance(f, dict):
            continue

        field_name = str(f.get("field_name", "")).strip()
        export_field_name = str(f.get("export_field_name", "")).strip()
        field_id = str(first(f.get("id"), f.get("export_field_id"), f.get("field_id"))).strip()

        if field_name == "" and export_field_name == "" and field_id == "":
            continue

        rows.append({
            "id": field_id,
            "field_name": field_name,
            "export_field_name": export_field_name
        })

    return rows


def extract_fields_json(exp):
    export_fields = exp.get("export_fields", [])
    if not isinstance(export_fields, list):
        export_fields = []
    return dumps_safe(export_fields)


def normalize_db_field_item(item):
    """
    Extrae field_name desde distintos formatos posibles del endpoint db_fields.
    """
    if isinstance(item, str):
        return item.strip(), item

    if not isinstance(item, dict):
        return "", item

    field_name = first(
        item.get("field_name"),
        item.get("name"),
        item.get("db_field"),
        item.get("value"),
        item.get("field")
    )

    return str(field_name).strip(), item


# =========================
# BLOCK BUILD
# =========================
def build_output_rows(exp=None, sch=None, error_message=""):
    if error_message:
        return [
            ["", "", "", ""],
            ["cron", "", "cron_timezone", ""],
            ["file_name", "", "destination", ""],
            ["export_protocol", "", "host", ""],
            ["username", "", "password", ""],
            ["login_type", "", "fetch_status", "ERROR"],
            ["error_message", str(error_message), "last_checked_at", now()],
            ["raw_export_json", "", "field_json", ""],
        ]

    rows = [
        ["", "", "", ""],
        ["cron", extract_cron(exp, sch), "cron_timezone", extract_cron_timezone(exp, sch)],
        ["file_name", first(exp.get("file_name"), exp.get("name")), "destination", first(exp.get("destination"), exp.get("ftp_path"), exp.get("path"))],
        ["export_protocol", first(exp.get("protocol"), exp.get("type"), exp.get("method")), "host", first(exp.get("host"), exp.get("ftp_host"), exp.get("server"))],
        ["username", first(exp.get("username"), exp.get("ftp_user"), exp.get("user")), "password", first(exp.get("password"), exp.get("ftp_password"))],
        ["login_type", first(exp.get("login_type"), exp.get("auth_type")), "fetch_status", "SUCCESS"],
        ["error_message", "", "last_checked_at", now()],
        ["raw_export_json", dumps_safe(exp), "field_json", extract_fields_json(exp)],
    ]

    field_rows = extract_field_rows(exp)

    if field_rows:
        rows.append(["", "", "", ""])
        for item in field_rows:
            rows.append([
                "id",
                item["id"],
                "Action",
                ""
            ])
            rows.append([
                "field_name",
                item["field_name"],
                "export_field_name",
                item["export_field_name"]
            ])

    return rows


# =========================
# WRITE
# =========================
def clear_block_output(service, start_col):
    start_letter = col_to_letter(start_col)
    end_letter = col_to_letter(start_col + 3)

    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{start_letter}3:{end_letter}400"
    ).execute()


def write_block_output(service, start_col, rows):
    start_letter = col_to_letter(start_col)
    end_letter = col_to_letter(start_col + 3)
    end_row = 2 + len(rows)

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!{start_letter}3:{end_letter}{end_row}",
        valueInputOption="RAW",
        body={"values": rows}
    ).execute()


def write_available_fields_sheet(service, rows):
    ensure_sheet_exists(service, AVAILABLE_FIELDS_SHEET)

    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{AVAILABLE_FIELDS_SHEET}!A:Z"
    ).execute()

    headers = [
        "db_name",
        "db_id",
        "field_name",
        "raw_field_json"
    ]

    values = [headers]
    for row in rows:
        values.append([
            row.get("db_name", ""),
            row.get("db_id", ""),
            row.get("field_name", ""),
            row.get("raw_field_json", "")
        ])

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{AVAILABLE_FIELDS_SHEET}!A1",
        valueInputOption="RAW",
        body={"values": values}
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

    available_fields_rows = []
    seen_db_ids = set()

    for block in blocks:
        start_col = block["start_col"]
        db_name = block["db_name"]
        db_id_raw = block["db_id"]
        export_id_raw = block["export_id"]
        export_name = block["export_name"]

        try:
            clear_block_output(service, start_col)

            if not db_id_raw or not export_id_raw:
                raise Exception("db_id y export_id son requeridos")

            db_id = clean_int(db_id_raw)
            export_id = clean_int(export_id_raw)

            # Update Exports block
            exports = get_exports(db_id)
            exp = find_export(exports, export_id=export_id, export_name=export_name)

            if not exp:
                raise Exception("Export not found")

            sch = get_schedule_optional(db_id, export_id)
            output_rows = build_output_rows(exp=exp, sch=sch)

            write_block_output(service, start_col, output_rows)
            print(f"✅ OK | col {start_col} | {db_name} | export_id {export_id}")

            # Available Fields from DATABASE
            if db_id not in seen_db_ids:
                seen_db_ids.add(db_id)

                db_fields_payload, err = get_db_fields(db_id)
                if err:
                    print(f"⚠️ Error leyendo db_fields para db_id {db_id}: {err}")
                else:
                    for item in db_fields_payload:
                        field_name, raw_item = normalize_db_field_item(item)
                        if field_name == "":
                            continue

                        available_fields_rows.append({
                            "db_name": db_name,
                            "db_id": str(db_id),
                            "field_name": field_name,
                            "raw_field_json": dumps_safe(raw_item)
                        })

        except Exception as e:
            error_rows = build_output_rows(error_message=str(e))
            write_block_output(service, start_col, error_rows)
            print(f"❌ ERROR | col {start_col} | {db_name} | {str(e)}")

            # Aunque falle el export, intenta traer db_fields si hay db_id
            try:
                if db_id_raw:
                    db_id = clean_int(db_id_raw)
                    if db_id not in seen_db_ids:
                        seen_db_ids.add(db_id)

                        db_fields_payload, err = get_db_fields(db_id)
                        if not err:
                            for item in db_fields_payload:
                                field_name, raw_item = normalize_db_field_item(item)
                                if field_name == "":
                                    continue

                                available_fields_rows.append({
                                    "db_name": db_name,
                                    "db_id": str(db_id),
                                    "field_name": field_name,
                                    "raw_field_json": dumps_safe(raw_item)
                                })
            except Exception:
                pass

    # dedupe por db_id + field_name
    deduped = []
    seen = set()
    for row in available_fields_rows:
        key = (row["db_id"], row["field_name"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)

    deduped.sort(key=lambda x: (x["db_name"], x["field_name"]))

    write_available_fields_sheet(service, deduped)
    print(f"✅ Available Fields actualizado con {len(deduped)} rows")


if __name__ == "__main__":
    main()
