import os
from datetime import datetime
from zoneinfo import ZoneInfo
import requests
import pandas as pd
import urllib3
from concurrent.futures import ThreadPoolExecutor, as_completed
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# FEEDONOMICS API
# =========================
API_KEY = os.environ["FEEDONOMICS_API_KEY"]
SERVICE_PATH = "https://meta.feedonomics.com/api.php"

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "x-api-key": API_KEY,
    "Content-Type": "application/json"
}

# =========================
# GOOGLE SHEETS
# =========================
SERVICE_ACCOUNT_FILE = "merchantapi-fdx.json"
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw")
DATABASES_SHEET = os.getenv("DATABASES_SHEET", "DataBases")
OUTPUT_SHEET = os.getenv("OUTPUT_SHEET", "Global Imports")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_WORKERS = int(os.getenv("MAX_WORKERS", "10"))

# =========================
# GOOGLE SHEETS CLIENT
# =========================
def get_sheets_service():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
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

def clean_text(value):
    if value is None:
        return None
    if not isinstance(value, str):
        value = str(value)

    value = value.replace("\ufeff", "")
    value = value.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    value = " ".join(value.split())
    return value.strip()

def get_last_updated_cst() -> str:
    return datetime.now(ZoneInfo("America/Chicago")).strftime("%Y-%m-%d %H:%M:%S %Z")

# =========================
# READ DATABASES FROM SHEET
# =========================
def read_databases_from_sheet(spreadsheet_id: str, sheet_name: str):
    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:Z"
    ).execute()

    values = result.get("values", [])
    if not values:
        raise Exception(f"La hoja '{sheet_name}' está vacía.")

    headers_row = values[0]
    data_rows = values[1:]

    df = pd.DataFrame(data_rows, columns=headers_row)

    df.columns = [str(c).strip().upper() for c in df.columns]

    required_cols = {"DB_NAME", "DB_ID"}
    missing = required_cols - set(df.columns)
    if missing:
        raise Exception(f"Faltan columnas requeridas en '{sheet_name}': {missing}")

    df = df[["DB_NAME", "DB_ID"]].copy()
    df = df[df["DB_ID"].astype(str).str.strip() != ""].copy()

    df["DB_ID"] = df["DB_ID"].astype(str).str.strip()
    df["DB_NAME"] = df["DB_NAME"].astype(str).str.strip()

    databases = []
    for _, row in df.iterrows():
        try:
            db_id = int(float(row["DB_ID"]))
        except Exception:
            continue

        databases.append({
            "db_name": row["DB_NAME"],
            "db_id": db_id
        })

    if not databases:
        raise Exception(f"No se encontraron databases válidos en '{sheet_name}'.")

    return databases

# =========================
# ENSURE SHEET EXISTS
# =========================
def ensure_sheet_exists(spreadsheet_id: str, sheet_name: str):
    service = get_sheets_service()

    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]

    if sheet_name not in sheets:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {"addSheet": {"properties": {"title": sheet_name}}}
                ]
            }
        ).execute()

# =========================
# FULL REFRESH TO SHEET
# =========================
def write_df_to_sheet_full_refresh(df: pd.DataFrame, spreadsheet_id: str, sheet_name: str):
    service = get_sheets_service()

    ensure_sheet_exists(spreadsheet_id, sheet_name)

    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=sheet_name
    ).execute()

    if df.empty:
        values = [[
            "db_name", "db_id", "import_id", "import_name", "import_field", "mapped_field",
            "import_file_name", "import_host", "import_password", "import_protocol",
            "import_tracker_field", "import_url", "import_username", "last_updated"
        ]]
    else:
        df = df.fillna("")
        values = [df.columns.tolist()] + df.astype(str).values.tolist()

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": values}
    ).execute()

# =========================
# FEEDONOMICS API
# =========================
def get_imports(db_id: int):
    url = f"{SERVICE_PATH}/dbs/{db_id}/imports"
    resp = requests.get(url, headers=HEADERS, verify=False, timeout=60)

    if resp.status_code != 200:
        return None, (resp.status_code, resp.text)

    payload = safe_json(resp)

    if isinstance(payload, dict) and payload.get("status") == "fail":
        return None, (resp.status_code, str(payload))

    if not isinstance(payload, list):
        return None, (resp.status_code, f"Unexpected payload: {type(payload)} | {payload}")

    return payload, None

# =========================
# WORKER
# =========================
def fetch_import_mappings_task(db_id: int, db_name: str, last_updated: str):
    try:
        imports_, err = get_imports(db_id)

        if err:
            return {
                "ok": False,
                "db_id": db_id,
                "db_name": db_name,
                "status": err[0],
                "error": err[1]
            }

        if not imports_:
            return {
                "ok": True,
                "rows": []
            }

        rows_local = []

        for imp in imports_:
            import_id = imp.get("id")
            import_name = imp.get("name")

            file_map = imp.get("file_map") or {}
            maps = file_map.get("maps") or {}

            import_info = imp.get("import_info") or {}

            import_file_name = import_info.get("file_name")
            import_host = import_info.get("host")
            import_password = import_info.get("password")
            import_protocol = import_info.get("protocol")
            import_tracker_field = import_info.get("tracker_field")
            import_url = import_info.get("url")
            import_username = import_info.get("username")

            if isinstance(maps, dict) and maps:
                for import_field, mapped_field in maps.items():
                    rows_local.append({
                        "db_name": db_name,
                        "db_id": db_id,
                        "import_id": import_id,
                        "import_name": import_name,
                        "import_field": clean_text(import_field),
                        "mapped_field": clean_text(mapped_field),
                        "import_file_name": import_file_name,
                        "import_host": import_host,
                        "import_password": import_password,
                        "import_protocol": import_protocol,
                        "import_tracker_field": import_tracker_field,
                        "import_url": import_url,
                        "import_username": import_username,
                        "last_updated": last_updated
                    })
            else:
                rows_local.append({
                    "db_name": db_name,
                    "db_id": db_id,
                    "import_id": import_id,
                    "import_name": import_name,
                    "import_field": None,
                    "mapped_field": None,
                    "import_file_name": import_file_name,
                    "import_host": import_host,
                    "import_password": import_password,
                    "import_protocol": import_protocol,
                    "import_tracker_field": import_tracker_field,
                    "import_url": import_url,
                    "import_username": import_username,
                    "last_updated": last_updated
                })

        return {
            "ok": True,
            "rows": rows_local
        }

    except Exception as ex:
        return {
            "ok": False,
            "db_id": db_id,
            "db_name": db_name,
            "status": "exception",
            "error": str(ex)
        }

# =========================
# MAIN
# =========================
def main():
    last_updated = get_last_updated_cst()

    print("📥 Leyendo databases desde Google Sheet...")
    databases = read_databases_from_sheet(SPREADSHEET_ID, DATABASES_SHEET)
    print(f"✅ Databases encontrados: {len(databases)}")

    mapping_rows = []
    errores = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(
                fetch_import_mappings_task,
                db["db_id"],
                db["db_name"],
                last_updated
            )
            for db in databases
        ]

        for future in as_completed(futures):
            result = future.result()

            if not result.get("ok"):
                errores.append({
                    "db_id": result.get("db_id"),
                    "db_name": result.get("db_name"),
                    "step": "imports",
                    "status": result.get("status"),
                    "error": result.get("error"),
                    "last_updated": last_updated
                })
                print(f"❌ Error imports DB {result.get('db_id')} ({result.get('db_name')})")
                continue

            mapping_rows.extend(result.get("rows", []))
            print(f"✅ Imports OK DB {result.get('db_id')} ({result.get('db_name')})")

    df = pd.DataFrame(mapping_rows)

    if not df.empty:
        df = df.sort_values(
            by=["db_name", "import_name", "import_field"],
            na_position="last"
        ).reset_index(drop=True)

    print("🧹 Haciendo refresh completo de la tab Global Imports...")
    write_df_to_sheet_full_refresh(df, SPREADSHEET_ID, OUTPUT_SHEET)

    print("✅ Datos guardados correctamente en Google Sheets")
    print(f"📄 Spreadsheet ID: {SPREADSHEET_ID}")
    print(f"📑 Source tab: {DATABASES_SHEET}")
    print(f"📑 Output tab: {OUTPUT_SHEET}")
    print(f"🔎 Total mappings encontrados: {len(df)}")
    print(f"🕒 Last updated: {last_updated}")

    if errores:
        df_errors = pd.DataFrame(errores)
        print(f"⚠️ Errores encontrados: {len(df_errors)}")
        print(df_errors.head(10).to_string(index=False))

if __name__ == "__main__":
    main()
