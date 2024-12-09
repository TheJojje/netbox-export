import requests
import openpyxl
from datetime import datetime

# Konfiguration
NETBOX_URL = "https://din-netbox-url/api/"
API_TOKEN = "din-api-token"
HEADERS = {"Authorization": f"Token {API_TOKEN}"}

# Funktion för att hämta subnät från NetBox
def fetch_prefixes():
    response = requests.get(f"{NETBOX_URL}ipam/prefixes/", headers=HEADERS)
    response.raise_for_status()
    return response.json()["results"]

# Funktion för att hämta IP-adresser för ett specifikt subnät
def fetch_ips(prefix_id):
    response = requests.get(f"{NETBOX_URL}ipam/ip-addresses/?parent={prefix_id}", headers=HEADERS)
    response.raise_for_status()
    return response.json()["results"]

# Skapa Excel-fil
def create_excel_with_subnets():
    # Hämta subnät
    prefixes = fetch_prefixes()
    
    # Skapa Excel-fil
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)  # Ta bort standardbladet

    for prefix in prefixes:
        sheet_name = prefix["prefix"].replace("/", "_")
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]  # Begränsa till 31 tecken

        # Skapa flik för subnät
        sheet = workbook.create_sheet(title=sheet_name)
        sheet.append(["IP Address", "Description", "Status", "Assigned To"])

        # Hämta IP-adresser för subnätet
        ips = fetch_ips(prefix["id"])
        for ip in ips:
            sheet.append([
                ip["address"],
                ip.get("description", ""),
                ip["status"]["label"],
                ip.get("assigned_object", {}).get("name", "")
            ])

    # Spara Excel-fil med dagens datum i filnamnet
    date_str = datetime.now().strftime("%Y-%m-%d")
    file_name = f"netbox_export_{date_str}.xlsx"
    workbook.save(file_name)
    print(f"Fil sparad: {file_name}")

# Kör skriptet
if __name__ == "__main__":
    create_excel_with_subnets()