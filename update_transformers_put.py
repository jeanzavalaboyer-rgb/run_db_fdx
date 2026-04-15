import os
import requests
import urllib3
import json
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# CONFIG
# =========================
api_key = os.environ["FEEDONOMICS_API_KEY"]
service_path = os.getenv("FEEDONOMICS_SERVICE_PATH", "https://meta.feedonomics.com/api.php")

headers = {
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Content-Type": "application/json"
}

SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "merchantapi-fdx.json")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "165va_Om_aFEmHg7h_zOUKpafsxEdB6MKvAycu6w16yw")
SHEET_NAME = os.getenv("SHEET_NAME", "Update Transformers")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# =========================
# GOOGLE SHEETS
# =========================
def get_sheets_service():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)

def read_sheet():
    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A:J"
    ).execute()

    values = result.get("values", [])

    if not values:
        raise Exception(f"No se encontraron datos en la tab '{SHEET_NAME}'")

    headers_row = values[0]
    data_rows = values[1:]

    normalized_rows = []
    for row in data_rows:
        row = row + [""] * (len(headers_row) - len(row))
        normalized_rows.append(row[:len(headers_row)])

    df = pd.DataFrame(normalized_rows, columns=headers_row)
    df = df.replace("", pd.NA).dropna(how="all").fillna("")

    return df

# =========================
# HELPERS
# =========================
def clean_int(value):
    value = str(value).strip()
    if not value:
        raise ValueError("Valor entero vacío")
    return int(float(value))

def normalize_bool(value):
    value = str(value).strip().lower()
    if value in ["true", "1", "yes", "y"]:
        return True
    if value in ["false", "0", "no", "n"]:
        return False
    raise ValueError(f"Valor inválido para enabled: {value}")

def normalize_selector(value):
    value = str(value).strip()
    if not value:
        raise ValueError("New Selector vacío")

    lowered = value.lower()
    if lowered == "true":
        return "true"
    if lowered == "false":
        return "false"

    return value

def normalize_transformer(value):
    value = str(value).strip()

    if not value:
        raise ValueError("New Transformer vacío")

    # Si parece una expresión de Feedonomics, dejar tal cual
    expression_signs = ["(", ")", "[", "]", ","]
    if any(sign in value for sign in expression_signs):
        return value

    # Si ya viene correctamente entre comillas simples
    if value.startswith("'") and value.endswith("'"):
        return value

    # Limpiar comillas sueltas y envolver como string literal
    clean_value = value.replace("'", "").strip()
    return f"'{clean_value}'"

def parse_exports(exports_value):
    if exports_value is None:
        return [0]

    exports_str = str(exports_value).strip()
    if not exports_str:
        return [0]

    try:
        parsed = json.loads(exports_str)
        export_ids = parsed.get("export_ids", [0])

        if isinstance(export_ids, list):
            normalized = []
            for item in export_ids:
                try:
                    normalized.append(int(item))
                except Exception:
                    normalized.append(int(float(str(item))))
            return normalized

        return [int(export_ids)]
    except Exception:
        return [0]

# =========================
# FEEDONOMICS UPDATE
# =========================
def update_transformer(row):
    db_id = clean_int(row["db_id"])
    transformer_id = clean_int(row["transformer_id"])
    field_name = str(row["field_name"]).strip()

    if not field_name:
        raise ValueError("field_name vacío")

    selector = normalize_selector(row["New Selector"])
    transformer = normalize_transformer(row["New Transformer"])
    enabled = normalize_bool(row["enabled"])
    export_ids = parse_exports(row["exports"])

    url = f"{service_path}/dbs/{db_id}/transformers/{transformer_id}"

    payload = {
        "enabled": enabled,
        "field_name": field_name,
        "selector": selector,
        "transformer": transformer,
        "export_id": export_ids
    }

    resp = requests.put(
        url,
        headers=headers,
        json=payload,
        verify=False,
        timeout=60
    )

    try:
        response_text = json.dumps(resp.json(), ensure_ascii=False)
    except Exception:
        response_text = resp.text

    return resp.status_code, response_text, payload

# =========================
# WRITE STATUS IN K:L
# =========================
def write_status(results):
    service = get_sheets_service()

    values = [["update_status", "error_message"]]

    for r in results:
        values.append([r["status"], r["error"]])

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!K1",
        valueInputOption="RAW",
        body={"values": values}
    ).execute()

# =========================
# MAIN
# =========================
def main():
    df = read_sheet()

    required_columns = [
        "db_name",
        "db_id",
        "field_name",
        "transformer_id",
        "selector",
        "transformer",
        "enabled",
        "exports",
        "New Selector",
        "New Transformer"
    ]

    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        raise Exception(f"Faltan columnas requeridas: {missing_cols}")

    results = []

    for idx, row in df.iterrows():
        sheet_row = idx + 2

        try:
            db_id_raw = str(row.get("db_id", "")).strip()
            transformer_id_raw = str(row.get("transformer_id", "")).strip()

            if not db_id_raw and not transformer_id_raw:
                status = "SKIPPED"
                error = "Fila omitida: db_id y transformer_id vacíos"
                print(f"⏭️ Fila {sheet_row} omitida")
            else:
                status_code, response, payload = update_transformer(row)

                if status_code == 200:
                    status = "SUCCESS"
                    error = ""
                    print(f"✅ Fila {sheet_row} OK")
                else:
                    status = "ERROR"
                    error = response
                    print(f"❌ Fila {sheet_row} ERROR {status_code}")

                print("Payload enviado:")
                print(json.dumps(payload, indent=2, ensure_ascii=False))
                print("Respuesta:")
                print(response)
                print("-" * 80)

        except Exception as e:
            status = "ERROR"
            error = str(e)
            print(f"❌ Fila {sheet_row} EXCEPTION: {error}")
            print("-" * 80)

        results.append({
            "status": status,
            "error": error
        })

    write_status(results)

    print("\n✅ Resultados escritos en columnas K y L")
    print("K = update_status")
    print("L = error_message")

if __name__ == "__main__":
    main()
