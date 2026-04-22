import os
import json
import requests
import urllib3
from concurrent.futures import ThreadPoolExecutor, as_completed
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

SPREADSHEET_ID = "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw"
SHEET_NAME = "Update Imports"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_WORKERS = 10

# OUTPUT desde columna F
BASE_HEADERS = [
    "import_id",
    "name",
    "cron",
    "cron_timezone",
    "cxn_id",
    "file_location",
    "import_info.file_name",
    "import_info.host",
    "import_info.password",
    "import_info.protocol",
    "import_info.username",
    "import_info.tracker_field",
    "import_info.url",
    "load_threshold",
    "load_threshold_percent",
    "update_threshold",
    "max_attempts",
    "stats.rows_loaded",
    "stats.rows_updated",
    "raw_import_json"
]

# =========================
# GOOGLE SHEETS
# =========================
def get_sheets_service():
    if not GOOGLE_SERVICE_ACCOUNT_JSON.strip():
        raise Exception("La variable de entorno GOOGLE_SERVICE_ACCOUNT_JSON no existe o está vacía.")

    service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def read_input_rows(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!D:E"
    ).execute()

    values = result.get("values", [])
    if not values:
        return []

    rows = []
    for i, row in enumerate(values[1:], start=2):
        db_name = row[0].strip() if len(row) > 0 else ""
        db_id = row[1].strip() if len(row) > 1 else ""

        if not db_name and not db_id:
            continue

        rows.append({
            "sheet_row": i,
            "db_name": db_name,
            "db_id": db_id
        })

    return rows


def clear_output_area(service):
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!B2:ZZ"
    ).execute()


def write_headers(service, output_headers):
    data = [
        {
            "range": f"{SHEET_NAME}!B1:C1",
            "values": [["status_message", "error_message"]]
        },
        {
            "range": f"{SHEET_NAME}!F1",
            "values": [output_headers]
        }
    ]

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={
            "valueInputOption": "RAW",
            "data": data
        }
    ).execute()


def write_full_rows(service, rows_matrix):
    if not rows_matrix:
        return

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!B2",
        valueInputOption="RAW",
        body={"values": rows_matrix}
    ).execute()


# =========================
# HELPERS
# =========================
def safe_json(resp):
    try:
        return resp.json()
    except Exception:
        return {"_raw": resp.text}


def clean_int(value):
    return int(float(str(value).strip()))


def clean_text(value):
    if value is None:
        return ""
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def extract_list_payload(payload):
    if isinstance(payload, list):
        return payload

    if isinstance(payload, dict):
        for key in ["imports", "data", "results", "rows"]:
            if isinstance(payload.get(key), list):
                return payload[key]

    return []


def pick_import_id(item):
    for key in ["id", "import_id"]:
        if item.get(key):
            return str(item[key]).strip()
    return ""


def get_nested(data, path, default=""):
    current = data
    for key in path.split("."):
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current is None:
            return default
    return current


def json_to_string(data):
    try:
        return json.dumps(data, ensure_ascii=False)
    except Exception:
        return str(data)


def get_maps(import_payload):
    file_map = import_payload.get("file_map", {})
    maps = file_map.get("maps", {})
    return maps if isinstance(maps, dict) else {}


def build_mapping_pairs(maps_dict):
    pairs = []
    for import_field, mapped_field in maps_dict.items():
        pairs.append(clean_text(mapped_field))
        pairs.append(clean_text(import_field))
    return pairs


def build_dynamic_headers(max_pairs):
    headers_out = BASE_HEADERS.copy()
    for i in range(1, max_pairs + 1):
        headers_out.append(f"mapped_field_{i}")
        headers_out.append(f"import_field_{i}")
    return headers_out


# =========================
# API
# =========================
def get_imports(db_id):
    url = f"{service_path}/dbs/{db_id}/imports"
    resp = requests.get(url, headers=headers, verify=False, timeout=60)

    payload = safe_json(resp)

    if resp.status_code != 200:
        raise Exception(payload)

    return extract_list_payload(payload)


# =========================
# WORKER
# =========================
def fetch_imports_for_input_row(row_item):
    db_name = row_item["db_name"]
    db_id_raw = row_item["db_id"]

    try:
        db_id = clean_int(db_id_raw)
        imports_list = get_imports(db_id)

        if not imports_list:
            return {
                "ok": True,
                "db_name": db_name,
                "db_id": db_id,
                "status": "SUCCESS",
                "error": "",
                "rows": [],
                "max_pairs": 0
            }

        output_rows = []
        max_pairs_local = 0

        for imp in imports_list:
            maps = get_maps(imp)
            mapping_pairs = build_mapping_pairs(maps)

            pair_count = len(mapping_pairs) // 2
            max_pairs_local = max(max_pairs_local, pair_count)

            row_values = [
                clean_text(pick_import_id(imp)),
                clean_text(imp.get("name")),
                clean_text(imp.get("cron")),
                clean_text(imp.get("cron_timezone")),
                clean_text(imp.get("cxn_id")),
                clean_text(imp.get("file_location")),
                clean_text(get_nested(imp, "import_info.file_name")),
                clean_text(get_nested(imp, "import_info.host")),
                clean_text(get_nested(imp, "import_info.password")),
                clean_text(get_nested(imp, "import_info.protocol")),
                clean_text(get_nested(imp, "import_info.username")),
                clean_text(get_nested(imp, "import_info.tracker_field")),
                clean_text(get_nested(imp, "import_info.url")),
                clean_text(imp.get("load_threshold")),
                clean_text(imp.get("load_threshold_percent")),
                clean_text(imp.get("update_threshold")),
                clean_text(imp.get("max_attempts")),
                clean_text(get_nested(imp, "stats.rows_loaded")),
                clean_text(get_nested(imp, "stats.rows_updated")),
                json_to_string(imp)
            ] + mapping_pairs

            output_rows.append(row_values)

        return {
            "ok": True,
            "db_name": db_name,
            "db_id": db_id,
            "status": "SUCCESS",
            "error": "",
            "rows": output_rows,
            "max_pairs": max_pairs_local
        }

    except Exception as e:
        return {
            "ok": False,
            "db_name": db_name,
            "db_id": db_id_raw,
            "status": "ERROR",
            "error": clean_text(str(e)),
            "rows": [],
            "max_pairs": 0
        }


# =========================
# MAIN
# =========================
def main():
    service = get_sheets_service()
    input_rows = read_input_rows(service)

    if not input_rows:
        print("No rows found")
        return

    indexed_inputs = list(enumerate(input_rows))
    results_by_index = {}
    max_pairs_global = 0

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_map = {
            executor.submit(fetch_imports_for_input_row, item): idx
            for idx, item in indexed_inputs
        }

        for future in as_completed(future_map):
            idx = future_map[future]
            result = future.result()
            results_by_index[idx] = result
            max_pairs_global = max(max_pairs_global, result["max_pairs"])

    output_headers = build_dynamic_headers(max_pairs_global)
    total_output_len = len(output_headers)

    final_rows = []

    for idx in range(len(indexed_inputs)):
        result = results_by_index[idx]

        if not result["rows"]:
            final_rows.append([
                result["status"],
                result["error"],
                result["db_name"],
                result["db_id"]
            ] + [""] * total_output_len)
            continue

        for i, row in enumerate(result["rows"]):
            padded = row + [""] * (total_output_len - len(row))

            final_rows.append([
                result["status"] if i == 0 else "",
                result["error"] if i == 0 else "",
                result["db_name"],
                result["db_id"]
            ] + padded)

    clear_output_area(service)
    write_headers(service, output_headers)
    write_full_rows(service, final_rows)

    print("✅ Done")


if __name__ == "__main__":
    main()
