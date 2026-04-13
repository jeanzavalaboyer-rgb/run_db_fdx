import os
from datetime import datetime
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
import pandas as pd
import urllib3
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
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
OUTPUT_SHEET = os.getenv("OUTPUT_SHEET", "Global Transformers")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Puedes ajustar estos números
DB_FIELDS_WORKERS = int(os.getenv("DB_FIELDS_WORKERS", "50"))
TRANSFORMERS_WORKERS = int(os.getenv("TRANSFORMERS_WORKERS", "50"))

# =========================
# SESSION HELPERS
# =========================
_thread_local = {}

def get_session():
    session = _thread_local.get("session")
    if session is None:
        session = requests.Session()

        retry = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=0.5,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"]
        )
        adapter = HTTPAdapter(
            max_retries=retry,
            pool_connections=100,
            pool_maxsize=100
        )
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        _thread_local["session"] = session

    return session

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

def get_last_updated_cst() -> str:
    return datetime.now(ZoneInfo("America/Chicago")).strftime("%Y-%m-%d %H:%M:%S %Z")

# =========================
# SHEETS HELPERS
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

    df["DB_NAME"] = df["DB_NAME"].astype(str).str.strip()
    df["DB_ID"] = df["DB_ID"].astype(str).str.strip()

    databases = []
    seen = set()

    for _, row in df.iterrows():
        try:
            db_id = int(float(row["DB_ID"]))
        except Exception:
            continue

        db_name = row["DB_NAME"]
        key = (db_name, db_id)

        if key in seen:
            continue
        seen.add(key)

        databases.append({
            "db_name": db_name,
            "db_id": db_id
        })

    if not databases:
        raise Exception(f"No se encontraron databases válidos en '{sheet_name}'.")

    return databases

def write_df_to_sheet_full_refresh(df: pd.DataFrame, spreadsheet_id: str, sheet_name: str):
    service = get_sheets_service()

    ensure_sheet_exists(spreadsheet_id, sheet_name)

    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=sheet_name
    ).execute()

    if df.empty:
        values = [[
            "db_name",
            "db_id",
            "field_name",
            "transformer_id",
            "export_id",
            "sort_order",
            "selector",
            "transformer",
            "enabled",
            "exports",
            "created_at",
            "last_updated"
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
def get_db_fields(db_id: int):
    session = get_session()
    url = f"{SERVICE_PATH}/dbs/{db_id}/db_fields"
    resp = session.get(url, headers=HEADERS, verify=False, timeout=(10, 60))

    if resp.status_code != 200:
        return None, (resp.status_code, resp.text)

    payload = safe_json(resp)

    if isinstance(payload, dict) and payload.get("status") == "fail":
        return None, (resp.status_code, str(payload))

    if not isinstance(payload, list):
        return None, (resp.status_code, f"Unexpected payload: {type(payload)} | {payload}")

    return payload, None

def get_transformers(db_id: int, field_name: str):
    session = get_session()
    url = f"{SERVICE_PATH}/dbs/{db_id}/transformers"
    params = {"field_name": field_name}

    resp = session.get(url, headers=HEADERS, params=params, verify=False, timeout=(10, 60))

    if resp.status_code != 200:
        return None, (resp.status_code, resp.text)

    payload = safe_json(resp)

    if isinstance(payload, dict) and payload.get("status") == "fail":
        return None, (resp.status_code, str(payload))

    if not isinstance(payload, list):
        return None, (resp.status_code, f"Unexpected payload: {type(payload)} | {payload}")

    return payload, None

# =========================
# WORKERS
# =========================
def fetch_db_fields_task(db: dict):
    db_id = db["db_id"]
    db_name = db["db_name"]

    try:
        fields, err = get_db_fields(db_id)

        if err:
            return {
                "ok": False,
                "db_id": db_id,
                "db_name": db_name,
                "step": "db_fields",
                "status": err[0],
                "error": err[1]
            }

        field_names = []
        for f in fields:
            field_name = f.get("field_name")
            if field_name:
                field_names.append(field_name)

        return {
            "ok": True,
            "db_id": db_id,
            "db_name": db_name,
            "field_names": field_names
        }

    except Exception as e:
        return {
            "ok": False,
            "db_id": db_id,
            "db_name": db_name,
            "step": "db_fields",
            "status": "exception",
            "error": str(e)
        }

def fetch_transformers_task(db_id: int, db_name: str, field_name: str, last_updated: str):
    try:
        transformers, err = get_transformers(db_id, field_name)

        if err:
            return {
                "ok": False,
                "db_id": db_id,
                "db_name": db_name,
                "field_name": field_name,
                "status": err[0],
                "error": err[1]
            }

        if not transformers:
            return {"ok": True, "rows": []}

        rows_local = []
        for t in transformers:
            rows_local.append({
                "db_name": db_name,
                "db_id": db_id,
                "field_name": field_name,
                "transformer_id": t.get("id"),
                "export_id": t.get("export_id"),
                "sort_order": t.get("sort_order"),
                "selector": t.get("selector"),
                "transformer": t.get("transformer"),
                "enabled": t.get("enabled"),
                "exports": t.get("exports"),
                "created_at": t.get("created_at"),
                "last_updated": last_updated
            })

        return {"ok": True, "rows": rows_local}

    except Exception as e:
        return {
            "ok": False,
            "db_id": db_id,
            "db_name": db_name,
            "field_name": field_name,
            "status": "exception",
            "error": str(e)
        }

# =========================
# MAIN
# =========================
def main():
    last_updated = get_last_updated_cst()

    print("📥 Leyendo databases desde Google Sheet...")
    databases = read_databases_from_sheet(SPREADSHEET_ID, DATABASES_SHEET)
    print(f"✅ Databases encontrados: {len(databases)}")

    rows = []
    errores = []

    print(f"🚀 Buscando db_fields en paralelo | workers: {DB_FIELDS_WORKERS}")
    db_field_tasks = []

    with ThreadPoolExecutor(max_workers=DB_FIELDS_WORKERS) as executor:
        futures = [executor.submit(fetch_db_fields_task, db) for db in databases]

        for future in as_completed(futures):
            result = future.result()

            if not result.get("ok"):
                errores.append({
                    "db_id": result.get("db_id"),
                    "db_name": result.get("db_name"),
                    "step": result.get("step"),
                    "status": result.get("status"),
                    "error": result.get("error"),
                    "last_updated": last_updated
                })
                print(f"❌ DB {result.get('db_id')} ({result.get('db_name')}) db_fields error")
                continue

            db_id = result["db_id"]
            db_name = result["db_name"]
            field_names = result.get("field_names", [])

            print(f"✅ DB {db_id} ({db_name}) | fields: {len(field_names)}")

            for field_name in field_names:
                db_field_tasks.append((db_id, db_name, field_name))

    print(f"🚀 Buscando transformers en paralelo | tasks: {len(db_field_tasks)} | workers: {TRANSFORMERS_WORKERS}")

    with ThreadPoolExecutor(max_workers=TRANSFORMERS_WORKERS) as executor:
        futures = [
            executor.submit(fetch_transformers_task, db_id, db_name, field_name, last_updated)
            for db_id, db_name, field_name in db_field_tasks
        ]

        for future in as_completed(futures):
            result = future.result()

            if not result.get("ok"):
                errores.append({
                    "db_id": result.get("db_id"),
                    "db_name": result.get("db_name"),
                    "step": f"transformers:{result.get('field_name')}",
                    "status": result.get("status"),
                    "error": result.get("error"),
                    "last_updated": last_updated
                })
                continue

            rows.extend(result.get("rows", []))

    df = pd.DataFrame(rows)

    if not df.empty:
        if "sort_order" in df.columns:
            df["sort_order"] = pd.to_numeric(df["sort_order"], errors="coerce")

        df = df.sort_values(
            by=["db_name", "field_name", "export_id", "sort_order"],
            na_position="last"
        ).reset_index(drop=True)

    print("🧹 Haciendo refresh completo de la tab Global Transformers...")
    write_df_to_sheet_full_refresh(df, SPREADSHEET_ID, OUTPUT_SHEET)

    print("✅ Datos guardados correctamente en Google Sheets")
    print(f"📄 Spreadsheet ID: {SPREADSHEET_ID}")
    print(f"📑 Source tab: {DATABASES_SHEET}")
    print(f"📑 Output tab: {OUTPUT_SHEET}")
    print(f"🔎 Total transformers encontrados: {len(df)}")
    print(f"🕒 Last updated: {last_updated}")

    if errores:
        df_errors = pd.DataFrame(errores)
        print(f"⚠️ Errores encontrados: {len(df_errors)}")
        print(df_errors.head(10).to_string(index=False))

if __name__ == "__main__":
    main()
