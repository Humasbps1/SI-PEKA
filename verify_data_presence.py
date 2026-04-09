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
match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
key = match.group(1)
sh = client.open_by_key(key)

print(f"Spreadsheet Title: {sh.title}")
print(f"Spreadsheet Key: {key}")

# Global search for "Imunisasi Campak"
print("\nSearching for 'Imunisasi Campak' in all worksheets...")
for ws in sh.worksheets():
    try:
        cell = ws.find("Imunisasi Campak")
        if cell:
            print(f"FOUND in worksheet '{ws.title}' at row {cell.row}, col {cell.col}")
            # Print the whole row
            row_data = ws.row_values(cell.row)
            print(f"Row data: {row_data}")
    except gspread.exceptions.CellNotFound:
        pass
    except Exception as e:
        print(f"Error in '{ws.title}': {e}")
