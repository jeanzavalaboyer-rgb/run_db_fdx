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
service_path = os.getenv("FEEDONOMICS_SERVICE_PATH", "https://meta.feedonomics.com/api.php")

headers = {
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Content-Type": "application/json"
}

# =========================
# GOOGLE SHEETS
# =========================
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "merchantapi-fdx.json")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw")

DATABASES_SHEET = os.getenv("DATABASES_SHEET", "Global Databases")
TARGET_SHEET = os.getenv("TARGET_SHEET", "Update Transformers")

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


def ensure_sheet_exists(spreadsheet_id: str, sheet_name: str):
    service = get_sheets_service()

    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

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


def read_sheet_raw(spreadsheet_id: str, sheet_name: str, range_a1: str = "A:Z"):
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
    padded = [r + [""] * (max_cols - len(r)) for r in values]
    return padded


# =========================
# READ GLOBAL DATABASES
# =========================
def read_global_databases(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    values = read_sheet_raw(spreadsheet_id, sheet_name, "A:Z")
    if not values:
        raise Exception(f"La hoja '{sheet_name}' está vacía.")

    values = pad_rows(values)

    headers_row = [str(h).strip().upper() for h in values[0]]
    data_rows = values[1:]

    df = pd.DataFrame(data_rows, columns=headers_row)

    # eliminar columnas duplicadas por si acaso
    df = df.loc[:, ~df.columns.duplicated()]

    required_cols = {"DB_NAME", "DB_ID"}
    missing = required_cols - set(df.columns)
    if missing:
        raise Exception(f"Faltan columnas requeridas en '{sheet_name}': {missing}")

    df = df[["DB_NAME", "DB_ID"]].copy()

    df["DB_NAME"] = df["DB_NAME"].astype(str).str.strip()
    df["DB_ID"] = df["DB_ID"].astype(str).str.strip()

    df = df[
        (df["DB_NAME"] != "") &
        (df["DB_ID"] != "")
    ].drop_duplicates().reset_index(drop=True)

    return df


# =========================
# READ UPDATE TRANSFORMERS
# =========================
def read_update_transformers(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """
    Lee A:C y fuerza estructura:
    A = DB_NAME
    B = DB_ID
    C = FIELD_NAME
    """
    values = read_sheet_raw(spreadsheet_id, sheet_name, "A:C")
    if not values:
        raise Exception(f"La hoja '{sheet_name}' está vacía.")

    values = pad_rows(values)

    # ignoramos headers originales y forzamos estructura por posición
    data_rows = values[1:] if len(values) > 1 else []

    df = pd.DataFrame(data_rows, columns=["DB_NAME", "DB_ID", "FIELD_NAME"])

    df["DB_NAME"] = df["DB_NAME"].astype(str).str.strip()
    df["DB_ID"] = df["DB_ID"].astype(str).str.strip()
    df["FIELD_NAME"] = df["FIELD_NAME"].astype(str).str.strip()

    return df


# =========================
# FEEDONOMICS API
# =========================
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


# =========================
# WRITE HELPERS
# =========================
def clear_range(spreadsheet_id: str, range_a1: str):
    service = get_sheets_service()
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=range_a1
    ).execute()


def write_single_column(spreadsheet_id: str, sheet_name: str, start_col: str, header: str, values_list: list):
    service = get_sheets_service()

    body_values = [[header]] + [[v] for v in values_list]

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{start_col}1",
        valueInputOption="RAW",
        body={"values": body_values}
    ).execute()


def write_table(spreadsheet_id: str, sheet_name: str, start_cell: str, headers_list: list, rows_list: list):
    service = get_sheets_service()

    body_values = [headers_list] + rows_list

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{start_cell}",
        valueInputOption="RAW",
        body={"values": body_values}
    ).execute()


# =========================
# MAIN
# =========================
def main():
    print("📥 Leyendo Global Databases...")
    df_global_db = read_global_databases(SPREADSHEET_ID, DATABASES_SHEET)

    print("📥 Leyendo Update Transformers...")
    df_update = read_update_transformers(SPREADSHEET_ID, TARGET_SHEET)

    # -------------------------
    # MAP DB_NAME -> DB_ID
    # -------------------------
    # Si hubiera db_name duplicado con distinto db_id, toma el primero
    db_map = (
        df_global_db
        .drop_duplicates(subset=["DB_NAME"], keep="first")
        .set_index("DB_NAME")["DB_ID"]
        .to_dict()
    )

    # llenar DB_ID en base a DB_NAME (col A)
    df_update["DB_ID"] = df_update["DB_NAME"].map(db_map).fillna("")

    print("✍️ Actualizando columna B con DB_ID...")
    write_single_column(
        spreadsheet_id=SPREADSHEET_ID,
        sheet_name=TARGET_SHEET,
        start_col="B",
        header="db_id",
        values_list=df_update["DB_ID"].tolist()
    )

    # -------------------------
    # DBs seleccionados en Update Transformers
    # -------------------------
    selected_dbs = (
        df_update[["DB_NAME", "DB_ID"]]
        .copy()
        .drop_duplicates()
    )

    selected_dbs = selected_dbs[
        (selected_dbs["DB_NAME"].astype(str).str.strip() != "") &
        (selected_dbs["DB_ID"].astype(str).str.strip() != "")
    ].copy()

    print(f"🔎 DBs seleccionados en Update Transformers: {len(selected_dbs)}")

    # -------------------------
    # Traer field_name vía API
    # -------------------------
    unique_rows = []
    seen = set()
    errors = []

    # cache para no repetir llamadas al mismo db_id
    db_fields_cache = {}

    for _, row in selected_dbs.iterrows():
        db_name = str(row["DB_NAME"]).strip()
        db_id_raw = str(row["DB_ID"]).strip()

        try:
            db_id = int(float(db_id_raw))
        except Exception:
            errors.append({
                "db_name": db_name,
                "db_id": db_id_raw,
                "error": "DB_ID inválido"
            })
            continue

        if db_id in db_fields_cache:
            fields, err = db_fields_cache[db_id]
        else:
            print(f"🚀 Consultando db_fields para DB {db_id} | {db_name}")
            fields, err = get_db_fields(db_id)
            db_fields_cache[db_id] = (fields, err)

        if err:
            errors.append({
                "db_name": db_name,
                "db_id": db_id,
                "error": f"status={err[0]} | {err[1]}"
            })
            continue

        for f in fields:
            field_name = str(f.get("field_name", "")).strip()
            if not field_name:
                continue

            key = (db_name, str(db_id), field_name)
            if key in seen:
                continue

            seen.add(key)
            unique_rows.append([
                db_name,
                str(db_id),
                field_name
            ])

    # ordenar por db_name y field_name
    if unique_rows:
        df_unique = pd.DataFrame(unique_rows, columns=["DB_NAME", "DB_ID", "FIELD_NAME"])
        df_unique = df_unique.sort_values(
            by=["DB_NAME", "FIELD_NAME"],
            na_position="last"
        ).reset_index(drop=True)

        output_rows = df_unique[["DB_NAME", "DB_ID", "FIELD_NAME"]].values.tolist()
    else:
        output_rows = []

    # -------------------------
    # Limpiar y escribir O:Q
    # -------------------------
    print("🧹 Limpiando columnas O:Q...")
    clear_range(SPREADSHEET_ID, f"{TARGET_SHEET}!O:Q")

    print("✍️ Escribiendo lista única en columnas O:Q...")
    write_table(
        spreadsheet_id=SPREADSHEET_ID,
        sheet_name=TARGET_SHEET,
        start_cell="O1",
        headers_list=["DB_NAME", "DB_ID", "FIELD_NAME"],
        rows_list=output_rows
    )

    print("✅ Proceso completado")
    print(f"   - DB_IDs actualizados en columna B: {len(df_update)} filas")
    print(f"   - Valores únicos escritos en O:Q: {len(output_rows)} filas")

    if errors:
        print("⚠️ Errores encontrados:")
        for e in errors[:20]:
            print(e)


if __name__ == "__main__":
    main()
