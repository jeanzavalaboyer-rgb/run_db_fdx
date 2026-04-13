import os
from datetime import datetime
import requests
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# FEEDONOMICS API
# =========================
api_key = os.environ["FEEDONOMICS_API_KEY"]
account_id = 1717
service_path = "https://meta.feedonomics.com/api.php"

headers = {
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Content-Type": "application/json"
}

# =========================
# API CALL FEEDONOMICS
# =========================
def get_account_status(account_id: int):
    url = f"{service_path}/accounts/{account_id}/status"
    resp = requests.get(url, headers=headers, verify=False, timeout=60)

    if resp.status_code != 200:
        raise Exception(f"Error {resp.status_code}: {resp.text}")

    return resp.json()

# =========================
# MAIN
# =========================
try:
    data = get_account_status(account_id)

    rows = []
    for db in data:
        rows.append({
            "db_id": db.get("id"),
            "db_name": db.get("name"),
            "data_count": db.get("data_count"),
            "import_status": db.get("import_status"),
            "export_status": db.get("export_status"),
            "num_imports": len(db.get("imports", [])),
            "num_exports": len(db.get("exports", []))
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.sort_values(
            by=["export_status", "import_status", "db_name"]
        ).reset_index(drop=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"global_databases_{timestamp}.xlsx"

    df.to_excel(output_file, index=False, sheet_name="Global Databases")

    print(f"✅ Excel generado correctamente: {output_file}")

    print("\n📊 Resumen export_status:")
    print(df["export_status"].value_counts(dropna=False))

    print("\n📊 Resumen import_status:")
    print(df["import_status"].value_counts(dropna=False))

except Exception as e:
    print(f"❌ Error: {str(e)}")
    raise
