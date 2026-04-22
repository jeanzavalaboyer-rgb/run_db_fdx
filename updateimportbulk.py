import os
import json
import requests
import urllib3
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

SPREADSHEET_ID = "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw"
SHEET_NAME = "Update Imports"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

TIMEOUT = 30


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
    raw_json["file_map"]["maps"] = raw_json["file_map"].get("maps") or {}
    raw_json["file_map"]["ignore_join"] = raw_json["file_map"].get("ignore_join") or []

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


def extract_sheet_maps(sheet_row):
    maps = {}
    idx = 1

    while True:
        mapped_key = f"mapped_field_{idx}"
        import_key = f"import_field_{idx}"

        exists_mapped = mapped_key in sheet_row
        exists_import = import_key in sheet_row

        if not exists_mapped and not exists_import:
            break

        mapped_val = clean(sheet_row.get(mapped_key))
        import_val = clean(sheet_row.get(import_key))

        if mapped_val == "" and import_val == "":
            idx += 1
            continue

        if mapped_val != "" and import_val != "":
            maps[import_val] = mapped_val

        idx += 1

    return maps


def detect_file_map_changes(sheet_row, raw_json):
    raw_maps = raw_json.get("file_map", {}).get("maps", {}) or {}
    sheet_maps = extract_sheet_maps(sheet_row)

    if not sheet_maps:
        return False

    normalized_raw = {clean(k): clean(v) for k, v in raw_maps.items()}
    normalized_sheet = {clean(k): clean(v) for k, v in sheet_maps.items()}

    return normalized_raw != normalized_sheet


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
    maps_from_sheet = extract_sheet_maps(sheet_row)

    if not maps_from_sheet:
        return None

    name_based_maps = normalize_flag(
        raw_json.get("name_based_maps", raw_file_map.get("name_based_maps", "0")),
        fallback="0"
    )

    encode_source_file_keys = to_int_or_default(
        raw_json.get("encode_source_file_keys", raw_file_map.get("encode_source_file_keys", 0)),
        default=0
    )

    clean_file_headers = normalize_flag(
        raw_json.get("clean_file_headers", raw_file_map.get("clean_file_headers", "0")),
        fallback="0"
    )

    file_type = clean(raw_json.get("file_type", raw_file_map.get("file_type", "delimited"))) or "delimited"

    final_maps = {}

    if name_based_maps == "1":
        for import_field, mapped_field in maps_from_sheet.items():
            encoded_key = text_to_hex(import_field)
            final_maps[encoded_key] = mapped_field
    else:
        raw_reverse = {}
        for raw_key, raw_mapped in raw_maps.items():
            raw_reverse[clean(raw_key)] = clean(raw_key)

        for import_field, mapped_field in maps_from_sheet.items():
            original_key = raw_reverse.get(clean(import_field), clean(import_field))
            final_maps[original_key] = mapped_field

    payload = {
        "encoding": clean(raw_file_map.get("encoding", "")),
        "separator": clean(raw_file_map.get("separator", ",")),
        "enclosure": clean(raw_file_map.get("enclosure", '"')),
        "escaper": clean(raw_file_map.get("escaper", '"')),
        "maps": final_maps,
        "ignore_join": raw_file_map.get("ignore_join", []),
        "name_based_maps": name_based_maps,
        "encode_source_file_keys": encode_source_file_keys,
        "clean_file_headers": clean_file_headers,
        "file_type": file_type
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


# =====================================================
# MAIN
# =====================================================
def main():
    service = get_service()
    headers, rows = read_sheet(service)

    if not rows:
        print("No rows found.")
        return

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

                write_status(
                    service,
                    idx,
                    "DELETED",
                    ""
                )

                print(f"🗑️ DELETED row {idx} | import_id {import_id}")
                continue

            raw_json = parse_raw(row_data.get("raw_import_json"))

            main_changes = detect_main_changes(row_data, raw_json)
            threshold_changes = detect_threshold_changes(row_data, raw_json)
            file_map_changes = detect_file_map_changes(row_data, raw_json)

            if not main_changes and not threshold_changes and not file_map_changes:
                write_status(
                    service,
                    idx,
                    "SKIPPED",
                    "No changes detected"
                )
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
                    raise Exception("File map payload could not be built")

                print(f"🔎 File map payload row {idx}: {payload_map}")

                resp = update_file_map(db_id, import_id, payload_map)

                if resp.status_code not in [200, 201]:
                    data = safe_json(resp)
                    raise Exception(
                        f"File map update failed | HTTP {resp.status_code} | {data}"
                    )

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

            write_status(
                service,
                idx,
                status_msg,
                detail_msg
            )

            print(f"✅ {status_msg} row {idx}")

        except Exception as e:
            msg = clean(str(e))
            write_status(service, idx, "ERROR", msg)
            print(f"❌ ERROR row {idx} | {msg}")


if __name__ == "__main__":
    main()
