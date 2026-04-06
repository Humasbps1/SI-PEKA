import gspread
from google.oauth2.service_account import Credentials
import toml
import re
import os

def test_connection():
    try:
        secrets_path = os.path.join(".streamlit", "secrets.toml")
        with open(secrets_path, "r") as f:
            secrets = toml.load(f)

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
        sheet_key = re.search(r'/d/([a-zA-Z0-9-_]+)', url).group(1)
        
        print(f"Connecting to: {url}")
        sh = client.open_by_key(sheet_key)
        print(f"✅ SUCCESS! Connected to Google Sheet: '{sh.title}'")
        
        worksheets = sh.worksheets()
        print(f"Found {len(worksheets)} worksheets: {', '.join([ws.title for ws in worksheets])}")
            
    except Exception as e:
        print(f"❌ CONNECTION FAILED: {e}")

if __name__ == "__main__":
    test_connection()
