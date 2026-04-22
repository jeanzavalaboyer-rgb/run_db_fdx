import os
import json
import requests
import urllib3
from concurrent.futures import ThreadPoolExecutor, as_completed
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =====================================================
# CONFIG
# =====================================================
API_KEY = os.environ["FEEDONOMICS_API_KEY"]
BASE_URL = "https://meta.feedonomics.com/api.php"
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "x-api-key": API_KEY,
    "Content-Type": "application/json",
    "Accept": "application/json"
}

session = requests.Session()
session.headers.update(HEADERS)

SPREADSHEET_ID = "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw"
INPUT_SHEET = "Update Imports"
EXPORT_DEP_SHEET = "Export Dependency"
TRANSFORMER_DEP_SHEET = "Transformer Dependency"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

TIMEOUT = 25
MAX_WORKERS = 20

LIGHT_RED = {
    "red": 0.98,
    "green": 0.88,
    "blue": 0.88
}


# =====================================================
# GOOGLE SHEETS
# =====================================================
def get_service():
    if not GOOGLE_SERVICE_ACCOUNT_JSON.strip():
        raise Exception("La variable de entorno GOOGLE_SERVICE_ACCOUNT_JSON no existe o está vacía.")

    service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def read_sheet(service, sheet_name):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:ZZ"
    ).execute()

    values = result.get("values", [])
    if not values:
        return [], []

    return values[0], values[1:]


def write_status_batch(service, updates):
    if not updates:
        return

    data = []
    for row_number, status_msg, error_msg in updates:
        data.append({
            "range": f"{INPUT_SHEET}!B{row_number}:C{row_number}",
            "values": [[status_msg, error_msg]]
        })

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={
            "valueInputOption": "RAW",
            "data": data
        }
    ).execute()


def clear_sheet(service, sheet_name):
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:ZZZ"
    ).execute()


def write_table(service, sheet_name, headers, rows):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": [headers] + rows}
    ).execute()


def ensure_sheet_exists(service, sheet_name):
    metadata = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID
    ).execute()

    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId")

    body = {
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

    resp = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body
    ).execute()

    replies = resp.get("replies", [])
    if replies:
        return replies[0]["addSheet"]["properties"]["sheetId"]

    raise Exception(f"Could not create sheet: {sheet_name}")


def get_sheet_id(service, sheet_name):
    metadata = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID
    ).execute()

    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId")

    raise Exception(f"Sheet not found: {sheet_name}")


def apply_cell_backgrounds(service, sheet_id, highlight_cells):
    if not highlight_cells:
        return

    requests_body = []
    for row_idx, col_idx in sorted(set(highlight_cells)):
        requests_body.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 1,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": LIGHT_RED
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        })

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests_body}
    ).execute()


# =====================================================
# HELPERS
# =====================================================
def clean(v):
    if v is None:
        return ""
    return str(v).strip()


def safe_json(resp):
    try:
        return resp.json()
    except Exception:
        return {"_raw": resp.text}


def row_to_dict(headers, row):
    data = {}
    for i, h in enumerate(headers):
        data[h] = row[i] if i < len(row) else ""
    return data


def dumps_safe(obj):
    try:
        return json.dumps(obj, ensure_ascii=False)
    except Exception:
        return str(obj)


def clean_int(v):
    return int(float(str(v).strip()))


def first(*vals):
    for v in vals:
        if v is not None and str(v).strip() != "":
            return v
    return ""


def should_process_row(row_data):
    action = clean(row_data.get("Action")).upper()
    status_message = clean(row_data.get("status_message")).upper()
    error_message = clean(row_data.get("error_message"))

    return (
        action == "GET DEPENDENCY"
        or "ERROR" in status_message
        or error_message != ""
    )


def extract_mapped_fields(row_data):
    mapped_fields = []

    for key, value in row_data.items():
        if str(key).startswith("mapped_field_"):
            v = clean(value)
            if v != "":
                mapped_fields.append(v)

    seen = set()
    final = []
    for x in mapped_fields:
        if x not in seen:
            seen.add(x)
            final.append(x)

    return final


def contains_field(text, field_name):
    text_s = clean(text)
    field_s = clean(field_name)

    if not text_s or not field_s:
        return False

    text_lower = text_s.lower()
    field_lower = field_s.lower()

    if f"[{field_lower}]" in text_lower:
        return True

    return field_lower in text_lower


# =====================================================
# API
# =====================================================
def get_exports(db_id):
    url = f"{BASE_URL}/dbs/{db_id}/exports"
    resp = session.get(url, verify=False, timeout=TIMEOUT)
    payload = safe_json(resp)

    if resp.status_code != 200:
        raise Exception(f"Exports HTTP {resp.status_code}: {dumps_safe(payload)}")
    if isinstance(payload, dict) and payload.get("status") == "fail":
        raise Exception(f"Exports fail payload: {dumps_safe(payload)}")
    if not isinstance(payload, list):
        raise Exception(f"Unexpected exports payload: {payload}")

    return payload


def get_transformers(db_id):
    url = f"{BASE_URL}/dbs/{db_id}/transformers"
    resp = session.get(url, verify=False, timeout=TIMEOUT)
    payload = safe_json(resp)

    if resp.status_code != 200:
        raise Exception(f"Transformers HTTP {resp.status_code}: {dumps_safe(payload)}")
    if isinstance(payload, dict) and payload.get("status") == "fail":
        raise Exception(f"Transformers fail payload: {dumps_safe(payload)}")
    if not isinstance(payload, list):
        raise Exception(f"Unexpected transformers payload: {payload}")

    return payload


def get_schedule_optional(db_id, export_id):
    try:
        url = f"{BASE_URL}/dbs/{db_id}/exports/{export_id}/schedule"
        resp = session.get(url, verify=False, timeout=TIMEOUT)
        if resp.status_code == 200:
            payload = safe_json(resp)
            if isinstance(payload, dict):
                return payload
    except Exception:
        pass
    return {}


# =====================================================
# EXPORT HELPERS
# =====================================================
def extract_export_selector(exp):
    return first(
        exp.get("export_selector"),
        exp.get("selector"),
        exp.get("select_rule"),
        exp.get("filter"),
        exp.get("where")
    )


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
            clean(first(f.get("id"), f.get("export_field_id"), f.get("field_id"))),
            clean(f.get("field_name")),
            clean(f.get("export_field_name"))
        ])
    return triplets


def extract_fields_json(exp):
    export_fields = exp.get("export_fields", [])
    if not isinstance(export_fields, list):
        export_fields = []
    return dumps_safe(export_fields)


def build_dynamic_export_headers(max_match_triplets):
    headers = [
        "import_id",
        "import_name",
        "mapped_field",
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
        "threshold",
    ]

    for i in range(1, max_match_triplets + 1):
        headers.extend([
            f"id_{i}",
            f"field_name_{i}",
            f"export_field_name_{i}"
        ])

    return headers


# =====================================================
# COLLECT EXPORTS / TRANSFORMERS
# =====================================================
def collect_export_dependencies_for_import(input_row, mapped_fields, exports, schedules_cache):
    grouped = {}

    for mapped_field in mapped_fields:
        for exp in exports:
            export_id = clean(exp.get("id"))
            export_name = clean(exp.get("name"))
            export_selector = extract_export_selector(exp)
            export_fields = exp.get("export_fields", [])
            if not isinstance(export_fields, list):
                export_fields = []

            selector_match = contains_field(export_selector, mapped_field)
            matched_triplets = []

            for f in export_fields:
                if not isinstance(f, dict):
                    continue

                this_field_name = clean(f.get("field_name"))
                if this_field_name.lower() == clean(mapped_field).lower():
                    matched_triplets.append([
                        clean(first(f.get("id"), f.get("export_field_id"), f.get("field_id"))),
                        this_field_name,
                        clean(f.get("export_field_name"))
                    ])

            field_name_match = len(matched_triplets) > 0

            if not field_name_match and not selector_match:
                continue

            unique_key = (
                input_row["sheet_row"],
                input_row["db_id"],
                input_row["import_id"],
                mapped_field,
                export_id
            )

            grouped[unique_key] = {
                "import_id": input_row["import_id"],
                "import_name": input_row["import_name"],
                "mapped_field": mapped_field,
                "db_name": input_row["db_name"],
                "db_id": input_row["db_id"],
                "export_id": export_id,
                "export_name": export_name,
                "export_obj": exp,
                "selector_match": selector_match,
                "field_name_match": field_name_match,
                "matched_triplets": matched_triplets
            }

    results = []
    max_triplets = 0

    for item in grouped.values():
        export_id = item["export_id"]
        db_id = item["db_id"]
        cache_key = (db_id, export_id)

        if cache_key not in schedules_cache:
            schedules_cache[cache_key] = get_schedule_optional(db_id, export_id)

        sch = schedules_cache.get(cache_key, {})
        exp = item["export_obj"]
        triplets = item["matched_triplets"]
        max_triplets = max(max_triplets, len(triplets))

        results.append({
            "import_id": item["import_id"],
            "import_name": item["import_name"],
            "mapped_field": item["mapped_field"],
            "db_name": item["db_name"],
            "db_id": item["db_id"],
            "export_id": export_id,
            "export_name": item["export_name"],
            "cron": extract_cron(exp, sch),
            "file_name": clean(first(exp.get("file_name"), exp.get("name"))),
            "export_protocol": clean(first(exp.get("protocol"), exp.get("type"), exp.get("method"))),
            "username": clean(first(exp.get("username"), exp.get("ftp_user"), exp.get("user"))),
            "cron_timezone": extract_cron_timezone(exp, sch),
            "destination": clean(first(exp.get("destination"), exp.get("ftp_path"), exp.get("path"))),
            "host": clean(first(exp.get("host"), exp.get("ftp_host"), exp.get("server"))),
            "password": clean(first(exp.get("password"), exp.get("ftp_password"))),
            "raw_export_json": dumps_safe(exp),
            "field_json": extract_fields_json(exp),
            "export_selector": extract_export_selector(exp),
            "threshold": extract_threshold(exp),
            "matched_triplets": triplets,
            "highlight_field_name": item["field_name_match"],
            "highlight_export_selector": item["selector_match"]
        })

    return results, max_triplets


def collect_transformer_dependencies_for_import(input_row, mapped_fields, transformers):
    grouped = {}

    for mapped_field in mapped_fields:
        for t in transformers:
            transformer_id = clean(t.get("id"))
            field_name = clean(t.get("field_name"))
            selector = clean(t.get("selector"))
            transformer_logic = clean(t.get("transformer"))
            enabled = clean(t.get("enabled"))
            exports = dumps_safe(t.get("exports"))

            field_name_match = field_name.lower() == clean(mapped_field).lower()
            selector_match = contains_field(selector, mapped_field)
            transformer_match = contains_field(transformer_logic, mapped_field)

            if not field_name_match and not selector_match and not transformer_match:
                continue

            unique_key = (
                input_row["sheet_row"],
                input_row["db_id"],
                input_row["import_id"],
                mapped_field,
                transformer_id
            )

            grouped[unique_key] = {
                "import_id": input_row["import_id"],
                "import_name": input_row["import_name"],
                "mapped_field": mapped_field,
                "db_name": input_row["db_name"],
                "db_id": input_row["db_id"],
                "field_name": field_name,
                "transformer_id": transformer_id,
                "selector": selector,
                "transformer": transformer_logic,
                "enabled": enabled,
                "exports": exports,
                "highlight_field_name": field_name_match,
                "highlight_selector": selector_match,
                "highlight_transformer": transformer_match
            }

    return list(grouped.values())


# =====================================================
# DB PREFETCH
# =====================================================
def prefetch_db_resources(db_id):
    with ThreadPoolExecutor(max_workers=2) as executor:
        future_exports = executor.submit(get_exports, db_id)
        future_transformers = executor.submit(get_transformers, db_id)

        exports = future_exports.result()
        transformers = future_transformers.result()

    return {
        "db_id": db_id,
        "exports": exports,
        "transformers": transformers,
        "error": None
    }


# =====================================================
# MAIN
# =====================================================
def main():
    service = get_service()

    headers, rows = read_sheet(service, INPUT_SHEET)
    if not rows:
        print("No rows found.")
        return

    ensure_sheet_exists(service, EXPORT_DEP_SHEET)
    ensure_sheet_exists(service, TRANSFORMER_DEP_SHEET)

    export_sheet_id = get_sheet_id(service, EXPORT_DEP_SHEET)
    transformer_sheet_id = get_sheet_id(service, TRANSFORMER_DEP_SHEET)

    input_items = []
    needed_db_ids = set()

    for idx, row in enumerate(rows, start=2):
        row_data = row_to_dict(headers, row)

        if not should_process_row(row_data):
            continue

        db_name = clean(row_data.get("db_name"))
        db_id_raw = clean(row_data.get("db_id"))
        import_id = clean(row_data.get("import_id"))
        import_name = clean(row_data.get("name")) or clean(row_data.get("import_name"))
        mapped_fields = extract_mapped_fields(row_data)

        input_items.append({
            "sheet_row": idx,
            "db_name": db_name,
            "db_id_raw": db_id_raw,
            "import_id": import_id,
            "import_name": import_name,
            "mapped_fields": mapped_fields
        })

        if db_id_raw != "":
            try:
                needed_db_ids.add(clean_int(db_id_raw))
            except Exception:
                pass

    exports_cache = {}
    transformers_cache = {}
    db_errors = {}
    schedules_cache = {}

    print(f"🚀 Prefetching DB resources for {len(needed_db_ids)} DB(s)...")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_map = {
            executor.submit(prefetch_db_resources, db_id): db_id
            for db_id in needed_db_ids
        }

        for future in as_completed(future_map):
            db_id = future_map[future]
            try:
                result = future.result()
                exports_cache[db_id] = result["exports"]
                transformers_cache[db_id] = result["transformers"]
                print(f"✅ Prefetched DB {db_id}")
            except Exception as e:
                db_errors[db_id] = clean(str(e))
                print(f"❌ Prefetch error DB {db_id} | {db_errors[db_id]}")

    export_results_all = []
    transformer_results_all = []
    status_updates = []
    processed = 0
    max_export_match_triplets = 0

    for item in input_items:
        idx = item["sheet_row"]
        db_id_raw = item["db_id_raw"]

        try:
            if db_id_raw == "":
                raise Exception("Missing db_id")

            db_id = clean_int(db_id_raw)
            mapped_fields = item["mapped_fields"]

            if not mapped_fields:
                raise Exception("No mapped_field_X values found")

            if db_id in db_errors:
                raise Exception(db_errors[db_id])

            exports = exports_cache.get(db_id, [])
            transformers = transformers_cache.get(db_id, [])

            input_info = {
                "sheet_row": idx,
                "db_name": item["db_name"],
                "db_id": str(db_id),
                "import_id": item["import_id"],
                "import_name": item["import_name"]
            }

            export_results, local_max_triplets = collect_export_dependencies_for_import(
                input_info,
                mapped_fields,
                exports,
                schedules_cache
            )
            transformer_results = collect_transformer_dependencies_for_import(
                input_info,
                mapped_fields,
                transformers
            )

            max_export_match_triplets = max(max_export_match_triplets, local_max_triplets)
            export_results_all.extend(export_results)
            transformer_results_all.extend(transformer_results)

            status_updates.append((
                idx,
                "DEPENDENCIES READY",
                f"Export rows: {len(export_results)} | Transformer rows: {len(transformer_results)}"
            ))

            processed += 1
            print(
                f"✅ Row {idx} | import_id={item['import_id']} | "
                f"mapped_fields={len(mapped_fields)} | "
                f"export_rows={len(export_results)} | "
                f"transformer_rows={len(transformer_results)}"
            )

        except Exception as e:
            msg = clean(str(e))
            status_updates.append((idx, "DEPENDENCY ERROR", msg))
            print(f"❌ Row {idx} | {msg}")

    export_headers = build_dynamic_export_headers(max_export_match_triplets)

    transformer_headers = [
        "import_id",
        "import_name",
        "mapped_field",
        "db_name",
        "db_id",
        "field_name",
        "transformer_id",
        "selector",
        "transformer",
        "enabled",
        "exports"
    ]

    export_rows_final = []
    export_highlights = []
    export_output_row_index = 1

    for item in export_results_all:
        row = [
            item["import_id"],
            item["import_name"],
            item["mapped_field"],
            item["db_name"],
            item["db_id"],
            item["export_id"],
            item["export_name"],
            item["cron"],
            item["file_name"],
            item["export_protocol"],
            item["username"],
            item["cron_timezone"],
            item["destination"],
            item["host"],
            item["password"],
            item["raw_export_json"],
            item["field_json"],
            item["export_selector"],
            item["threshold"],
        ]

        flat_triplets = []
        for triplet in item["matched_triplets"]:
            flat_triplets.extend(triplet)

        expected_triplet_cells = max_export_match_triplets * 3
        flat_triplets += [""] * (expected_triplet_cells - len(flat_triplets))
        row += flat_triplets
        export_rows_final.append(row)

        if item["highlight_export_selector"]:
            export_highlights.append((export_output_row_index, 17))

        if item["highlight_field_name"]:
            for i in range(len(item["matched_triplets"])):
                field_name_col = 19 + (i * 3) + 1
                export_highlights.append((export_output_row_index, field_name_col))

        export_output_row_index += 1

    transformer_rows_final = []
    transformer_highlights = []
    transformer_output_row_index = 1

    for item in transformer_results_all:
        row = [
            item["import_id"],
            item["import_name"],
            item["mapped_field"],
            item["db_name"],
            item["db_id"],
            item["field_name"],
            item["transformer_id"],
            item["selector"],
            item["transformer"],
            item["enabled"],
            item["exports"]
        ]
        transformer_rows_final.append(row)

        if item["highlight_field_name"]:
            transformer_highlights.append((transformer_output_row_index, 5))
        if item["highlight_selector"]:
            transformer_highlights.append((transformer_output_row_index, 7))
        if item["highlight_transformer"]:
            transformer_highlights.append((transformer_output_row_index, 8))

        transformer_output_row_index += 1

    print("🧹 Clearing output sheets...")
    clear_sheet(service, EXPORT_DEP_SHEET)
    clear_sheet(service, TRANSFORMER_DEP_SHEET)

    print("✍️ Writing Export Dependency...")
    write_table(service, EXPORT_DEP_SHEET, export_headers, export_rows_final)

    print("✍️ Writing Transformer Dependency...")
    write_table(service, TRANSFORMER_DEP_SHEET, transformer_headers, transformer_rows_final)

    print("🎨 Applying highlights...")
    apply_cell_backgrounds(service, export_sheet_id, export_highlights)
    apply_cell_backgrounds(service, transformer_sheet_id, transformer_highlights)

    print("✍️ Updating statuses...")
    write_status_batch(service, status_updates)

    print("✅ Done")
    print(f"Rows processed: {processed}")
    print(f"Export dependency rows: {len(export_rows_final)}")
    print(f"Transformer dependency rows: {len(transformer_rows_final)}")
    print(f"Max matched export triplets: {max_export_match_triplets}")


if __name__ == "__main__":
    main()
