import gspread
from google.oauth2.service_account import Credentials
import toml
import re
from datetime import datetime

# Load credentials
secrets = toml.load(".streamlit/secrets.toml")
s = secrets["connections"]["gsheets"]
sa = s["service_account"]
pk = sa["private_key"].replace("\\n", "\n")

creds = Credentials.from_service_account_info({
    "type": sa["type"],
    "project_id": sa["project_id"],
    "private_key": pk,
    "client_email": sa["client_email"],
    "token_uri": sa["token_uri"],
}, scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])

client = gspread.authorize(creds)
url = s["spreadsheet"]
sh = client.open_by_url(url)

print(f"Spreadsheet Title: {sh.title}")
# Get metadata from Drive API if possible, or just check worksheets
print(f"Worksheets: {[ws.title for ws in sh.worksheets()]}")

ws = sh.worksheet("📊Promosi Statistik 2026")
# Check if there are other columns that might be dates
headers = ws.row_values(1)
print(f"Headers: {headers}")

# Get rows 7-12 (No 6-11 approx)
rows = ws.get_values("A7:E12")
print("\nRaw Values for rows 7-12 (Col A to E):")
for i, row in enumerate(rows):
    print(f"Row {i+7}: {row}")

# Check value render options
print("\nChecking with UNFORMATTED_VALUE:")
rows_unformatted = ws.get_values("A7:E12", value_render_option="UNFORMATTED_VALUE")
for i, row in enumerate(rows_unformatted):
    print(f"Row {i+7}: {row}")

# Check if there's a specific 'Imunisasi Campak' text now
cell = None
try:
    cell = ws.find("Imunisasi Campak")
    if cell:
        print(f"\nFOUND 'Imunisasi Campak' at Row {cell.row}")
        print(f"Full Row {cell.row}: {ws.row_values(cell.row)}")
except:
    print("\n'Imunisasi Campak' NOT FOUND using find()")
