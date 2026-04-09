import gspread
from google.oauth2.service_account import Credentials
import toml
import re

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
key = re.search(r'/d/([a-zA-Z0-9-_]+)', url).group(1)
sh = client.open_by_key(key)

ws_name = "📊Promosi Statistik 2026" # Using the exact name from prev output
try:
    ws = sh.worksheet(ws_name)
except:
    # Try finding by partial name
    for sheet in sh.worksheets():
        if "promosi" in sheet.title.lower():
            ws = sheet
            break

data = ws.get_all_values()
# Save as CSV to verify
import csv
with open("raw_promo_data.csv", "w", encoding="utf-8", newline="") as f:
    writer = csv.writer(f)
    writer.writerows(data)

print(f"Worksheet '{ws.title}' saved to raw_promo_data.csv")
print(f"Total rows: {len(data)}")
