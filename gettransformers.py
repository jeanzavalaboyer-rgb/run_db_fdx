import os
import requests
import pandas as pd
import urllib3
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# FEEDONOMICS API
# =========================
api_key = os.environ["FEEDONOMICS_API_KEY"]
service_path = os.getenv(
    "FEEDONOMICS_SERVICE_PATH",
    "https://meta.feedonomics.com/api.php"
)

session = requests.Session()

session.headers.update({
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Accept": "application/json"
})

# =========================
# GOOGLE SHEETS
# =========================
SERVICE_ACCOUNT_FILE = os.getenv(
    "GOOGLE_APPLICATION_CREDENTIALS",
    "merchantapi-fdx.json"
)

SPREADSHEET_ID = os.getenv(
    "SPREADSHEET_ID",
    "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw"
)

TARGET_SHEET = os.getenv(
    "TARGET_SHEET",
    "Update Transformers"
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


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


def read_sheet_raw(
    spreadsheet_id: str,
    sheet_name: str,
    range_a1: str
):
    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{range_a1}"
    ).execute()

    return result.get("values", [])


def pad_rows(values):
    if not values:
        return []

    max_cols = max(len(r) for r in values)

    return [
        r + [""] * (max_cols - len(r))
        for r in values
    ]


def clear_range(
    spreadsheet_id: str,
    range_a1: str
):
    service = get_sheets_service()

    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=range_a1
    ).execute()


def write_table(
    spreadsheet_id: str,
    sheet_name: str,
    start_cell: str,
    headers_list: list,
    rows_list: list
):
    service = get_sheets_service()

    body_values = [headers_list] + rows_list

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{start_cell}",
        valueInputOption="RAW",
        body={"values": body_values}
    ).execute()


# =========================
# READ UPDATE TRANSFORMERS
# =========================
def read_update_transformers_selected(
    spreadsheet_id: str,
    sheet_name: str
) -> pd.DataFrame:

    # AHORA LEE DESDE D:F
    # D = db_name
    # E = db_id
    # F = field_name

    values = read_sheet_raw(
        spreadsheet_id,
        sheet_name,
        "D:F"
    )

    if not values:
        raise Exception(
            f"La hoja '{sheet_name}' está vacía."
        )

    values = pad_rows(values)

    data_rows = values[1:] if len(values) > 1 else []

    df = pd.DataFrame(
        data_rows,
        columns=[
            "DB_NAME",
            "DB_ID",
            "FIELD_NAME"
        ]
    )

    df["DB_NAME"] = (
        df["DB_NAME"]
        .astype(str)
        .str.strip()
    )

    df["DB_ID"] = (
        df["DB_ID"]
        .astype(str)
        .str.strip()
    )

    df["FIELD_NAME"] = (
        df["FIELD_NAME"]
        .astype(str)
        .str.strip()
    )

    df = df[
        (df["DB_NAME"] != "") &
        (df["DB_ID"] != "") &
        (df["FIELD_NAME"] != "")
    ].copy()

    return (
        df
        .drop_duplicates()
        .reset_index(drop=True)
    )


# =========================
# FEEDONOMICS API
# =========================
def get_transformers(
    db_id: int,
    field_name: str
):
    url = f"{service_path}/dbs/{db_id}/transformers"

    params = {
        "field_name": field_name
    }

    resp = session.get(
        url,
        params=params,
        verify=False,
        timeout=60
    )

    if resp.status_code != 200:
        return None, (
            resp.status_code,
            resp.text
        )

    payload = safe_json(resp)

    if (
        isinstance(payload, dict)
        and payload.get("status") == "fail"
    ):
        return None, (
            resp.status_code,
            str(payload)
        )

    if not isinstance(payload, list):
        return None, (
            resp.status_code,
            f"Unexpected payload: {payload}"
        )

    return payload, None


# =========================
# MAIN
# =========================
def main():

    print("📥 Leyendo Update Transformers...")

    df_selected = read_update_transformers_selected(
        SPREADSHEET_ID,
        TARGET_SHEET
    )

    rows = []
    cache = {}

    for _, row in df_selected.iterrows():

        db_name = row["DB_NAME"]
        field_name = row["FIELD_NAME"]

        try:
            db_id = int(float(row["DB_ID"]))
        except:
            continue

        key = (db_id, field_name)

        if key not in cache:

            transformers, err = get_transformers(
                db_id,
                field_name
            )

            cache[key] = (transformers, err)

        else:
            transformers, err = cache[key]

        if err or not transformers:
            continue

        for t in transformers:

            rows.append([
                db_name,
                db_id,
                field_name,
                t.get("id"),
                t.get("selector"),
                t.get("transformer"),
                t.get("enabled"),
                t.get("exports")
            ])

    df_out = pd.DataFrame(rows)

    print("🧹 Limpiando columnas T:AA...")

    clear_range(
        SPREADSHEET_ID,
        f"{TARGET_SHEET}!T:AA"
    )

    print("✍️ Escribiendo resultados...")

    write_table(
        SPREADSHEET_ID,
        TARGET_SHEET,
        "T1",
        [
            "db_name",
            "db_id",
            "field_name",
            "transformer_id",
            "selector",
            "transformer",
            "enabled",
            "exports"
        ],
        df_out.values.tolist()
    )

    print("✅ Done")


if __name__ == "__main__":
    main()
