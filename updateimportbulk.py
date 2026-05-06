import os
import json
import requests
import urllib3
from requests.exceptions import ReadTimeout
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
    "Content-Type": "application/json"
}

SERVICE_ACCOUNT_FILE = "merchantapi-fdx.json"

SPREADSHEET_ID = "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw"
SHEET_NAME = "Update Imports"
MAPPED_FIELDS_SHEET = "Mapped Fields Output"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

TIMEOUT = 90
MAX_WORKERS = 8


# =====================================================
# GOOGLE SHEETS
# =====================================================
def get_service():
    if not GOOGLE_SERVICE_ACCOUNT_JSON or not GOOGLE_SERVICE_ACCOUNT_JSON.strip():
        raise Exception("La variable de entorno GOOGLE_SERVICE_ACCOUNT_JSON no existe o está vacía.")

    with open(SERVICE_ACCOUNT_FILE, "w", encoding="utf-8") as f:
        f.write(GOOGLE_SERVICE_ACCOUNT_JSON)

    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def read_sheet(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A:ZZ"
    ).execute()

    values = result.get("values", [])

    if not values:
        return [], []

    return values[0], values[1:]


def write_status(service, row_number, status_msg, error_msg):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!B{row_number}:C{row_number}",
        valueInputOption="RAW",
        body={"values": [[status_msg, error_msg]]}
    ).execute()


def get_spreadsheet_metadata(service):
    return service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID
    ).execute()


def get_sheet_id_by_name(metadata, sheet_name):
    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId")
    return None


def create_sheet_if_not_exists(service, sheet_name):
    metadata = get_spreadsheet_metadata(service)
    existing_id = get_sheet_id_by_name(metadata, sheet_name)

    if existing_id is not None:
        return existing_id

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

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body
    ).execute()

    metadata = get_spreadsheet_metadata(service)
    return get_sheet_id_by_name(metadata, sheet_name)


def clear_sheet(service, sheet_name):
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:ZZZ"
    ).execute()


def write_rows(service, sheet_name, rows):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": rows}
    ).execute()


# =====================================================
# HELPERS
# =====================================================
def clean(v):
    if v is None:
        return ""
    return str(v).strip()


def row_to_dict(headers, row):
    data = {}
    for i, h in enumerate(headers):
        data[h] = row[i] if i < len(row) else ""
    return data


def safe_json(resp):
    try:
        return resp.json()
    except Exception:
        return {"_raw": resp.text}


def parse_raw(raw_text):
    raw_text = clean(raw_text)

    if not raw_text:
        return {}

    raw_json = json.loads(raw_text)

    raw_json["import_info"] = raw_json.get("import_info") or {}
    raw_json["tags"] = raw_json.get("tags") or {"platform": "Not Applicable"}
    raw_json["file_map"] = raw_json.get("file_map") or {}

    if not isinstance(raw_json["file_map"], dict):
        raw_json["file_map"] = {}

    raw_json["file_map"]["maps"] = raw_json["file_map"].get("maps") or {}

    if not isinstance(raw_json["file_map"]["maps"], (dict, list)):
        raw_json["file_map"]["maps"] = {}

    raw_json["file_map"]["ignore_join"] = raw_json["file_map"].get("ignore_join") or []

    if not isinstance(raw_json["file_map"]["ignore_join"], list):
        raw_json["file_map"]["ignore_join"] = []

    return raw_json


def same(a, b):
    return clean(a) == clean(b)


def normalize_threshold_value(value, fallback="0"):
    val = clean(value)

    if val == "":
        return fallback

    try:
        num = float(val)
        if num.is_integer():
            return str(int(num))
        return str(num)
    except Exception:
        return val


def normalize_flag(value, fallback="0"):
    val = clean(value)

    if val == "":
        return str(fallback)

    if val.lower() in ["true", "yes", "y"]:
        return "1"

    if val.lower() in ["false", "no", "n"]:
        return "0"

    if val in ["0", "1"]:
        return val

    return str(fallback)


def to_int_or_default(value, default=0):
    val = clean(value)
    if val == "":
        return default
    try:
        return int(float(val))
    except Exception:
        return default


def text_to_hex(text):
    return clean(text).encode("utf-8").hex()


def hex_to_text_if_possible(value):
    value = clean(value)
    if value == "":
        return value

    try:
        if len(value) % 2 == 0 and all(c in "0123456789abcdefABCDEF" for c in value):
            return bytes.fromhex(value).decode("utf-8")
    except Exception:
        pass

    return value


def get_status_message(main_changes, threshold_changes, file_map_changes):
    parts = []

    if main_changes:
        parts.append("MAIN")
    if threshold_changes:
        parts.append("THRESHOLDS")
    if file_map_changes:
        parts.append("FILE_MAP")

    if not parts:
        return "SKIPPED"

    return " + ".join(parts) + " UPDATED"


def build_changed_fields_message(main_changes, threshold_changes, file_map_changes):
    changed = []
    changed.extend(sorted(main_changes.keys()))
    changed.extend(sorted(threshold_changes.keys()))

    if file_map_changes:
        changed.append("mapped_field_X/import_field_X")

    if not changed:
        return ""

    return "Changed: " + ", ".join(changed)


def unique_preserve_order(items):
    seen = set()
    out = []

    for item in items:
        val = clean(item)

        if val == "":
            continue

        if val not in seen:
            seen.add(val)
            out.append(val)

    return out


def extract_first_array_only(payload):
    if isinstance(payload, list):
        if payload and isinstance(payload[0], list):
            return payload[0]
        return payload

    if isinstance(payload, dict):
        for key in ["data", "rows", "results", "parsed", "file", "preview"]:
            value = payload.get(key)

            if isinstance(value, list):
                if value and isinstance(value[0], list):
                    return value[0]
                return value

    return []


def extract_fields_from_parsed_payload(payload):
    arr = extract_first_array_only(payload)

    if not isinstance(arr, list):
        return []

    decoded = [hex_to_text_if_possible(x) for x in arr]
    return unique_preserve_order(decoded)


# =====================================================
# CHANGE DETECTION
# =====================================================
def detect_main_changes(sheet_row, raw_json):
    changes = {}

    import_info = raw_json.get("import_info") or {}

    checks = [
        ("name", raw_json.get("name")),
        ("file_location", raw_json.get("file_location")),
        ("import_info.file_name", import_info.get("file_name")),
        ("import_info.host", import_info.get("host")),
        ("import_info.password", import_info.get("password")),
        ("import_info.protocol", import_info.get("protocol")),
        ("import_info.username", import_info.get("username")),
        ("cron", raw_json.get("cron")),
        ("cron_timezone", raw_json.get("cron_timezone")),
    ]

    for col, raw_val in checks:
        sheet_val = clean(sheet_row.get(col))

        if sheet_val == "":
            continue

        if not same(sheet_val, raw_val):
            changes[col] = sheet_val

    return changes


def detect_threshold_changes(sheet_row, raw_json):
    changes = {}

    checks = [
        ("load_threshold", raw_json.get("load_threshold")),
        ("load_threshold_percent", raw_json.get("load_threshold_percent")),
        ("update_threshold", raw_json.get("update_threshold")),
    ]

    for col, raw_val in checks:
        sheet_val = clean(sheet_row.get(col))

        if sheet_val == "":
            continue

        if not same(sheet_val, raw_val):
            changes[col] = sheet_val

    return changes


def extract_sheet_maps(sheet_row, max_consecutive_empty=10):
    maps = {}
    idx = 1
    consecutive_empty = 0

    while True:
        mapped_key = f"mapped_field_{idx}"
        import_key = f"import_field_{idx}"

        mapped_exists = mapped_key in sheet_row
        import_exists = import_key in sheet_row

        mapped_val = clean(sheet_row.get(mapped_key, ""))
        import_val = clean(sheet_row.get(import_key, ""))

        if not mapped_exists and not import_exists:
            consecutive_empty += 1
            if consecutive_empty >= max_consecutive_empty:
                break
            idx += 1
            continue

        if mapped_val == "" and import_val == "":
            consecutive_empty += 1
            if consecutive_empty >= max_consecutive_empty:
                break
            idx += 1
            continue

        consecutive_empty = 0

        if mapped_val != "" and import_val != "":
            maps[import_val] = mapped_val

        idx += 1

    return maps


def normalize_raw_maps_to_plain(raw_maps):
    normalized = {}

    if isinstance(raw_maps, dict):
        for raw_key, mapped_field in raw_maps.items():
            plain_key = hex_to_text_if_possible(raw_key)
            normalized[clean(plain_key)] = clean(mapped_field)

    elif isinstance(raw_maps, list):
        # Some Feedonomics imports return file_map.maps as a list.
        # We skip map comparison to avoid blocking threshold updates.
        return {}

    else:
        return {}

    return normalized


def extract_changed_sheet_maps(sheet_row, raw_json):
    raw_maps = raw_json.get("file_map", {}).get("maps", {}) or {}
    raw_plain_maps = normalize_raw_maps_to_plain(raw_maps)
    sheet_maps = extract_sheet_maps(sheet_row)

    changed = {}

    for import_field, mapped_field in sheet_maps.items():
        import_field_clean = clean(import_field)
        mapped_field_clean = clean(mapped_field)
        raw_mapped = clean(raw_plain_maps.get(import_field_clean))

        if raw_mapped == "" or raw_mapped != mapped_field_clean:
            changed[import_field_clean] = mapped_field_clean

    return changed


def build_merged_plain_maps(sheet_row, raw_json):
    raw_maps = raw_json.get("file_map", {}).get("maps", {}) or {}
    raw_plain_maps = normalize_raw_maps_to_plain(raw_maps)
    changed_sheet_maps = extract_changed_sheet_maps(sheet_row, raw_json)

    merged = dict(raw_plain_maps)

    for import_field, mapped_field in changed_sheet_maps.items():
        merged[import_field] = mapped_field

    return merged


def detect_file_map_changes(sheet_row, raw_json):
    raw_maps = raw_json.get("file_map", {}).get("maps", {}) or {}

    if isinstance(raw_maps, list):
        return False

    changed_sheet_maps = extract_changed_sheet_maps(sheet_row, raw_json)
    return len(changed_sheet_maps) > 0


# =====================================================
# PAYLOAD BUILDERS
# =====================================================
def build_main_payload(sheet_row, raw_json, changes):
    import_info = raw_json.get("import_info") or {}

    payload = {
        "name": changes.get("name", raw_json.get("name")),
        "file_location": changes.get("file_location", raw_json.get("file_location")),
        "join_type": raw_json.get("join_type", "product_feed"),
        "file_name": changes.get("import_info.file_name", import_info.get("file_name", "")),
        "host": changes.get("import_info.host", import_info.get("host", "")),
        "password": changes.get("import_info.password", import_info.get("password", "")),
        "protocol": changes.get("import_info.protocol", import_info.get("protocol", "")),
        "username": changes.get("import_info.username", import_info.get("username", "")),
        "timeout": raw_json.get("timeout", "900"),
        "tags": raw_json.get("tags") or {"platform": "Not Applicable"},
        "do_import": True,
        "do_notify": False,
        "notification_email": None
    }

    if "cron" in changes:
        payload["cron"] = changes["cron"] if changes["cron"] != "" else None

    if "cron_timezone" in changes:
        payload["cron_timezone"] = changes["cron_timezone"]

    return payload


def build_threshold_payload(sheet_row, raw_json):
    return {
        "load_threshold": normalize_threshold_value(
            sheet_row.get("load_threshold"),
            fallback=normalize_threshold_value(raw_json.get("load_threshold"), "0")
        ),
        "load_threshold_percent": normalize_threshold_value(
            sheet_row.get("load_threshold_percent"),
            fallback=normalize_threshold_value(raw_json.get("load_threshold_percent"), "0")
        ),
        "update_threshold": normalize_threshold_value(
            sheet_row.get("update_threshold"),
            fallback=normalize_threshold_value(raw_json.get("update_threshold"), "1")
        )
    }


def build_file_map_payload(sheet_row, raw_json):
    raw_file_map = raw_json.get("file_map", {}) or {}

    raw_maps = raw_file_map.get("maps", {}) or {}
    if isinstance(raw_maps, list):
        return None

    changed_sheet_maps = extract_changed_sheet_maps(sheet_row, raw_json)

    if not changed_sheet_maps:
        return None

    merged_plain_maps = build_merged_plain_maps(sheet_row, raw_json)

    name_based_maps = normalize_flag(
        raw_json.get("name_based_maps", raw_file_map.get("name_based_maps", "0")),
        fallback="0"
    )

    if name_based_maps == "1":
        encode_source_file_keys = 1
    else:
        encode_source_file_keys = to_int_or_default(
            raw_json.get("encode_source_file_keys", raw_file_map.get("encode_source_file_keys", 0)),
            default=0
        )

    clean_file_headers = normalize_flag(
        raw_json.get("clean_file_headers", raw_file_map.get("clean_file_headers", "0")),
        fallback="0"
    )

    file_type = clean(
        raw_json.get("file_type", raw_file_map.get("file_type", "delimited"))
    ) or "delimited"

    final_maps = {}

    if name_based_maps == "1":
        for import_field, mapped_field in merged_plain_maps.items():
            final_maps[text_to_hex(import_field)] = mapped_field
    else:
        for import_field, mapped_field in merged_plain_maps.items():
            final_maps[clean(import_field)] = mapped_field

    raw_ignore_join = raw_file_map.get("ignore_join", [])
    if not isinstance(raw_ignore_join, list):
        raw_ignore_join = []

    final_ignore_join = []
    seen_ignore = set()

    for value in raw_ignore_join:
        value_clean = clean(value)
        if value_clean != "" and value_clean not in seen_ignore:
            final_ignore_join.append(value_clean)
            seen_ignore.add(value_clean)

    for _, mapped_field in changed_sheet_maps.items():
        mapped_field_clean = clean(mapped_field)
        if mapped_field_clean != "" and mapped_field_clean not in seen_ignore:
            final_ignore_join.append(mapped_field_clean)
            seen_ignore.add(mapped_field_clean)

    valid_mapped_fields = {clean(v) for v in merged_plain_maps.values() if clean(v) != ""}
    final_ignore_join = [x for x in final_ignore_join if x in valid_mapped_fields]

    payload = {
        "encoding": clean(raw_file_map.get("encoding", "")),
        "separator": clean(raw_file_map.get("separator", ",")),
        "enclosure": clean(raw_file_map.get("enclosure", '"')),
        "escaper": clean(raw_file_map.get("escaper", '"')),
        "maps": final_maps,
        "ignore_join": final_ignore_join,
        "file_type": file_type,
        "clean_file_headers": clean_file_headers,
        "name_based_maps": name_based_maps,
        "encode_source_file_keys": encode_source_file_keys
    }

    return payload


# =====================================================
# API
# =====================================================
def update_main(db_id, import_id, payload):
    url = f"{BASE_URL}/dbs/{db_id}/imports/{import_id}"

    return requests.post(
        url,
        json=payload,
        headers=HEADERS,
        verify=False,
        timeout=TIMEOUT
    )


def update_thresholds(db_id, import_id, payload):
    url = f"{BASE_URL}/dbs/{db_id}/imports/{import_id}/thresholds"

    return requests.put(
        url,
        json=payload,
        headers=HEADERS,
        verify=False,
        timeout=TIMEOUT
    )


def update_file_map(db_id, import_id, payload):
    url = f"{BASE_URL}/dbs/{db_id}/imports/{import_id}/file_map"

    return requests.put(
        url,
        json=payload,
        headers=HEADERS,
        verify=False,
        timeout=TIMEOUT
    )


def delete_import(db_id, import_id):
    url = f"{BASE_URL}/dbs/{db_id}/imports/{import_id}"

    return requests.delete(
        url,
        headers=HEADERS,
        verify=False,
        timeout=TIMEOUT
    )


def get_fields_before_mapping(db_id, import_id):
    url = f"{BASE_URL}/dbs/{db_id}/imports/{import_id}/file?format=parsed&limit=4"

    resp = requests.get(
        url,
        headers=HEADERS,
        verify=False,
        timeout=TIMEOUT
    )

    payload = safe_json(resp)

    if resp.status_code != 200:
        raise Exception(
            f"GET_FIELDS failed | HTTP {resp.status_code} | {payload}"
        )

    return payload


# =====================================================
# GET MAPPED FIELDS OUTPUT
# =====================================================
def build_mapped_fields_headers(max_fields):
    headers = [
        "source_row",
        "db_name",
        "db_id",
        "import_id",
        "status_message",
        "error_message"
    ]

    for i in range(1, max_fields + 1):
        headers.append(f"mapped_field_{i}")
        headers.append(f"import_field_{i}")

    return headers


def build_mapped_fields_row(result, max_fields):
    row = [
        result["source_row"],
        result["db_name"],
        result["db_id"],
        result["import_id"],
        result["status"],
        result["error"]
    ]

    fields = result["fields"]

    for field in fields:
        row.append(field)
        row.append(field)

    remaining = max_fields - len(fields)

    for _ in range(remaining):
        row.append("")
        row.append("")

    return row


def process_get_mapped_fields_row(item):
    db_name = item["db_name"]
    db_id_raw = item["db_id"]
    import_id_raw = item["import_id"]
    source_row = item["source_row"]

    try:
        db_id = clean(db_id_raw)
        import_id = clean(import_id_raw)

        if not db_id or not import_id:
            raise Exception("Missing db_id/import_id")

        payload = get_fields_before_mapping(db_id, import_id)
        fields = extract_fields_from_parsed_payload(payload)

        return {
            "source_row": source_row,
            "db_name": db_name,
            "db_id": db_id,
            "import_id": import_id,
            "status": "SUCCESS",
            "error": "",
            "fields": fields
        }

    except Exception as e:
        return {
            "source_row": source_row,
            "db_name": db_name,
            "db_id": db_id_raw,
            "import_id": import_id_raw,
            "status": "ERROR",
            "error": clean(str(e)),
            "fields": []
        }


def process_get_mapped_fields(service, headers, rows):
    action_rows = []

    for idx, row in enumerate(rows, start=2):
        row_data = row_to_dict(headers, row)
        action = clean(row_data.get("Action")).upper()

        if action != "GET MAPPED FIELDS":
            continue

        db_id = clean(row_data.get("db_id"))
        import_id = clean(row_data.get("import_id"))

        if not db_id or not import_id:
            write_status(service, idx, "ERROR", "Missing db_id/import_id")
            continue

        action_rows.append({
            "source_row": idx,
            "db_name": clean(row_data.get("db_name")),
            "db_id": db_id,
            "import_id": import_id
        })

    if not action_rows:
        return

    create_sheet_if_not_exists(service, MAPPED_FIELDS_SHEET)

    results = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(process_get_mapped_fields_row, item) for item in action_rows]

        for future in as_completed(futures):
            results.append(future.result())

    results.sort(key=lambda x: x["source_row"])

    max_fields = max((len(r["fields"]) for r in results), default=0)
    output = [build_mapped_fields_headers(max_fields)]

    for r in results:
        output.append(build_mapped_fields_row(r, max_fields))

        if r["status"] == "SUCCESS":
            write_status(service, r["source_row"], "GET MAPPED FIELDS", f"{len(r['fields'])} fields found")
        else:
            write_status(service, r["source_row"], "ERROR", r["error"])

    clear_sheet(service, MAPPED_FIELDS_SHEET)
    write_rows(service, MAPPED_FIELDS_SHEET, output)


# =====================================================
# MAIN
# =====================================================
def main():
    service = get_service()
    headers, rows = read_sheet(service)

    if not rows:
        print("No rows found.")
        return

    process_get_mapped_fields(service, headers, rows)

    for idx, row in enumerate(rows, start=2):
        row_data = row_to_dict(headers, row)
        action = clean(row_data.get("Action")).upper()

        if action not in ["UPDATE", "DELETE"]:
            continue

        try:
            db_id = clean(row_data.get("db_id"))
            import_id = clean(row_data.get("import_id"))

            if not db_id or not import_id:
                raise Exception("Missing db_id/import_id")

            if action == "DELETE":
                resp = delete_import(db_id, import_id)

                if resp.status_code not in [200, 201]:
                    data = safe_json(resp)
                    raise Exception(
                        f"Delete failed | HTTP {resp.status_code} | {data}"
                    )

                write_status(service, idx, "DELETED", "")
                print(f"🗑️ DELETED row {idx} | import_id {import_id}")
                continue

            raw_json = parse_raw(row_data.get("raw_import_json"))

            main_changes = detect_main_changes(row_data, raw_json)
            threshold_changes = detect_threshold_changes(row_data, raw_json)

            try:
                file_map_changes = detect_file_map_changes(row_data, raw_json)
            except Exception as e:
                file_map_changes = False
                print(f"⚠️ File map check skipped row {idx}: {e}")

            if not main_changes and not threshold_changes and not file_map_changes:
                write_status(service, idx, "SKIPPED", "No changes detected")
                print(f"⏭️ SKIPPED row {idx}")
                continue

            if main_changes:
                payload_main = build_main_payload(row_data, raw_json, main_changes)
                print(f"🔎 Main payload row {idx}: {payload_main}")

                resp = update_main(db_id, import_id, payload_main)

                if resp.status_code not in [200, 201]:
                    data = safe_json(resp)
                    raise Exception(
                        f"Main update failed | HTTP {resp.status_code} | {data}"
                    )

            if threshold_changes:
                payload_thr = build_threshold_payload(row_data, raw_json)
                print(f"🔎 Threshold payload row {idx}: {payload_thr}")

                resp = update_thresholds(db_id, import_id, payload_thr)

                if resp.status_code not in [200, 201]:
                    data = safe_json(resp)
                    raise Exception(
                        f"Threshold update failed | HTTP {resp.status_code} | {data}"
                    )

            if file_map_changes:
                payload_map = build_file_map_payload(row_data, raw_json)

                if not payload_map:
                    print(f"⚠️ File map update skipped row {idx}: payload could not be built")
                else:
                    print(f"🔎 File map payload row {idx}: {payload_map}")

                    try:
                        resp = update_file_map(db_id, import_id, payload_map)

                        if resp.status_code not in [200, 201]:
                            data = safe_json(resp)
                            raise Exception(
                                f"File map update failed | HTTP {resp.status_code} | {data}"
                            )

                    except ReadTimeout:
                        write_status(
                            service,
                            idx,
                            "TIMEOUT - VERIFY",
                            "File map request timed out after sending. Please verify in Feedonomics if the update was applied."
                        )
                        print(f"⏱️ TIMEOUT row {idx} | file_map update may have been applied")
                        continue

            status_msg = get_status_message(
                main_changes,
                threshold_changes,
                file_map_changes
            )

            detail_msg = build_changed_fields_message(
                main_changes,
                threshold_changes,
                file_map_changes
            )

            write_status(service, idx, status_msg, detail_msg)
            print(f"✅ {status_msg} row {idx}")

        except Exception as e:
            msg = clean(str(e))
            write_status(service, idx, "ERROR", msg)
            print(f"❌ ERROR row {idx} | {msg}")


if __name__ == "__main__":
    main()
