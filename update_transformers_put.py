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
        range=f"{SHEET_NAME}!A:Z"
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

    expression_signs = ["(", ")", "[", "]", ","]
    if any(sign in value for sign in expression_signs):
        return value

    if value.startswith("'") and value.endswith("'"):
        return value

    clean_value = value.replace("'", "").strip()
    return f"'{clean_value}'"


def parse_exports_object(exports_value):
    default_value = {"export_ids": ["0"], "all_exports": True}

    if exports_value is None:
        return default_value

    exports_str = str(exports_value).strip()
    if not exports_str:
        return default_value

    try:
        parsed = json.loads(exports_str)

        if not isinstance(parsed, dict):
            return default_value

        raw_export_ids = parsed.get("export_ids", ["0"])
        raw_all_exports = parsed.get("all_exports", False)

        if not isinstance(raw_export_ids, list):
            raw_export_ids = [raw_export_ids]

        normalized_ids = []

        for item in raw_export_ids:
            item_str = str(item).strip()

            if item_str == "":
                continue

            try:
                if "." in item_str:
                    item_str = str(int(float(item_str)))
            except Exception:
                pass

            normalized_ids.append(item_str)

        return {
            "export_ids": normalized_ids,
            "all_exports": bool(raw_all_exports)
        }

    except Exception:
        return default_value


def parse_export_ids(exports_value):
    exports_obj = parse_exports_object(exports_value)

    if exports_obj["all_exports"] is True:
        return ["0"]

    return exports_obj["export_ids"]


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

    exports_obj = parse_exports_object(row["exports"])
    export_ids = parse_export_ids(row["exports"])

    url = f"{service_path}/dbs/{db_id}/transformers/{transformer_id}"

    payload = {
        "enabled": enabled,
        "field_name": field_name,
        "selector": selector,
        "transformer": transformer,
        "export_id": export_ids,
        "exports": exports_obj
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
# FEEDONOMICS CREATE
# =========================
def create_transformer(row):
    db_id = clean_int(row["db_id"])
    field_name = str(row["field_name"]).strip()

    if not field_name:
        raise ValueError("field_name vacío")

    selector = normalize_selector(row["New Selector"])
    transformer = normalize_transformer(row["New Transformer"])
    enabled = normalize_bool(row["enabled"])

    exports_obj = parse_exports_object(row["exports"])
    export_ids = parse_export_ids(row["exports"])

    url = f"{service_path}/dbs/{db_id}/transformers"

    payload = {
        "enabled": enabled,
        "field_name": field_name,
        "selector": selector,
        "transformer": transformer,
        "export_id": export_ids,
        "exports": exports_obj
    }

    resp = requests.post(
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
# FEEDONOMICS DELETE
# =========================
def delete_transformer(row):
    db_id = clean_int(row["db_id"])
    transformer_id = clean_int(row["transformer_id"])

    url = f"{service_path}/dbs/{db_id}/transformers/{transformer_id}"

    resp = requests.delete(
        url,
        headers=headers,
        verify=False,
        timeout=60
    )

    try:
        response_text = json.dumps(resp.json(), ensure_ascii=False)
    except Exception:
        response_text = resp.text

    payload = {
        "db_id": db_id,
        "transformer_id": transformer_id
    }

    return resp.status_code, response_text, payload


# =========================
# WRITE STATUS IN B:C
# =========================
def write_status(results):
    service = get_sheets_service()

    values = [["update_status", "error_message"]]

    for r in results:
        values.append([r["status"], r["error"]])

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!B1",
        valueInputOption="RAW",
        body={"values": values}
    ).execute()


# =========================
# MAIN
# =========================
def main():
    df = read_sheet()

    required_columns = [
        "action",
        "update_status",
        "error_message",
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
            action = str(row.get("action", "")).strip().lower()

            if action not in ["update", "new", "delete"]:
                status = "SKIPPED"
                error = f"Fila omitida: action '{action}' no soportado"
                print(f"⏭️ Fila {sheet_row} omitida | action: {row.get('action', '')}")

            else:
                db_id_raw = str(row.get("db_id", "")).strip()
                transformer_id_raw = str(row.get("transformer_id", "")).strip()

                if not db_id_raw:
                    status = "ERROR"
                    error = "db_id vacío"
                    print(f"❌ Fila {sheet_row} ERROR: {error}")

                elif action in ["update", "delete"] and not transformer_id_raw:
                    status = "ERROR"
                    error = "transformer_id vacío para action Update/Delete"
                    print(f"❌ Fila {sheet_row} ERROR: {error}")

                else:
                    print(f"Fila {sheet_row} | action: {action.upper()}")

                    if action in ["update", "new"]:
                        exports_obj = parse_exports_object(row.get("exports", ""))
                        export_ids = parse_export_ids(row.get("exports", ""))

                        print(f"Fila {sheet_row} | exports original: {row.get('exports', '')}")
                        print(f"Fila {sheet_row} | exports parseado: {json.dumps(exports_obj, ensure_ascii=False)}")
                        print(f"Fila {sheet_row} | export_id enviado: {export_ids}")

                    if action == "update":
                        status_code, response, payload = update_transformer(row)

                    elif action == "new":
                        status_code, response, payload = create_transformer(row)

                    elif action == "delete":
                        status_code, response, payload = delete_transformer(row)

                    if status_code in [200, 201, 204]:
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

    print("\n✅ Resultados escritos en columnas B y C")
    print("B = update_status")
    print("C = error_message")


if __name__ == "__main__":
    main()
