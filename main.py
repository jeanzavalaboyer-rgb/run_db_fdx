import os
import requests
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

api_key = os.environ["FEEDONOMICS_API_KEY"]
account_id = int(os.getenv("ACCOUNT_ID", "1717"))
output_file = os.getenv("OUTPUT_NAME", "global_databases.xlsx")

service_path = "https://meta.feedonomics.com/api.php"

headers = {
    "Authorization": f"Bearer {api_key}",
    "x-api-key": api_key,
    "Content-Type": "application/json"
}

def get_account_status(account_id: int):
    url = f"{service_path}/accounts/{account_id}/status"
    resp = requests.get(url, headers=headers, verify=False, timeout=60)

    if resp.status_code != 200:
        raise Exception(f"Error {resp.status_code}: {resp.text}")

    return resp.json()

def main():
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

    df.to_excel(output_file, index=False, sheet_name="Global Databases")

    print(f"✅ Excel generado correctamente: {output_file}")

    if not df.empty:
        print("\n📊 Resumen export_status:")
        print(df["export_status"].value_counts(dropna=False))

        print("\n📊 Resumen import_status:")
        print(df["import_status"].value_counts(dropna=False))

if __name__ == "__main__":
    main()
