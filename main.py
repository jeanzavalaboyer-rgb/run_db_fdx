import os
from datetime import datetime
import requests
import pandas as pd
import urllib3
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# FEEDONOMICS API
# =========================
API_KEY = os.environ["FEEDONOMICS_API_KEY"]
ACCOUNT_ID = int(os.getenv("ACCOUNT_ID", "1717"))
SERVICE_PATH = "https://meta.feedonomics.com/api.php"

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "x-api-key": API_KEY,
    "Content-Type": "application/json",
}

# =========================
# GOOGLE SHEETS
# =========================
SERVICE_ACCOUNT_FILE = "merchantapi-fdx.json"
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw")
SHEET_NAME = os.getenv("SHEET_NAME", "Global Databases")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# =========================
# FEEDONOMICS CALL
# =========================
def get_account_status(account_id: int):
    url = f"{SERVICE_PATH}/accounts/{account_id}/status"
    resp = requests.get(url, headers=HEADERS, verify=False, timeout=60)

    if resp.status_code != 200:
        raise Exception(f"Error {resp.status_code}: {resp.text}")

    return resp.json()

# =========================
# GOOGLE SHEETS CLIENT
# =========================
def get_sheets_service():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES,
    )
    return build("sheets", "v4", credentials=creds)

# =========================
# VALIDATE / CREATE TAB
# =========================
def ensure_sheet_exists(service, spreadsheet_id: str, sheet_name: str):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]

    if sheet_name not in sheets:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {"addSheet": {"properties": {"title": sheet_name}}}
                ]
            },
        ).execute()

# =========================
# WRITE DATAFRAME TO SHEET
# =========================
def write_df_to_sheet(df: pd.DataFrame, spreadsheet_id: str, sheet_name: str):
    service = get_sheets_service()

    ensure_sheet_exists(service, spreadsheet_id, sheet_name)

    # Limpiar toda la tab
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=sheet_name,
    ).execute()

    # Preparar data
    df = df.fillna("")
    values = [df.columns.tolist()] + df.astype(str).values.tolist()

    # Escribir desde A1
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

# =========================
# BUILD DATAFRAME
# =========================
def build_dataframe(data) -> pd.DataFrame:
    rows = []
    for db in data:
        rows.append({
            "db_id": db.get("id"),
            "db_name": db.get("name"),
            "data_count": db.get("data_count"),
            "import_status": db.get("import_status"),
            "export_status": db.get("export_status"),
            "num_imports": len(db.get("imports", [])),
            "num_exports": len(db.get("exports", [])),
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.sort_values(
            by=["export_status", "import_status", "db_name"]
        ).reset_index(drop=True)

    # Última columna con fecha y hora de actualización
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df["last_updated"] = timestamp

    return df

# =========================
# MAIN
# =========================
def main():
    data = get_account_status(ACCOUNT_ID)
    df = build_dataframe(data)

    write_df_to_sheet(df, SPREADSHEET_ID, SHEET_NAME)

    print("✅ Datos guardados correctamente en Google Sheets")
    print(f"📄 Spreadsheet ID: {SPREADSHEET_ID}")
    print(f"📑 Tab: {SHEET_NAME}")

    if not df.empty:
        print("\n📊 Resumen export_status:")
        print(df["export_status"].value_counts(dropna=False))

        print("\n📊 Resumen import_status:")
        print(df["import_status"].value_counts(dropna=False))

        print("\n🕒 Last updated:")
        print(df["last_updated"].iloc[0])

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        raise
