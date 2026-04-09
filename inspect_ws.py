import gspread
from google.oauth2.service_account import Credentials
import toml
import os
import re

def inspect():
    try:
        secrets = toml.load('.streamlit/secrets.toml')
        s = secrets['connections']['gsheets']
        sa = s['service_account']
        pk = sa['private_key'].replace('\\n', '\n')
        
        creds = Credentials.from_service_account_info({
            'type': sa['type'], 'project_id': sa['project_id'], 'private_key': pk, 
            'client_email': sa['client_email'], 'token_uri': sa['token_uri']
        }, scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'])
        
        client = gspread.authorize(creds)
        url = s['spreadsheet']
        sheet_key = re.search(r'/d/([a-zA-Z0-9-_]+)', url).group(1)
        sh = client.open_by_key(sheet_key)
        
        skip_sheets = [
            'pilih', 'config', 'referensi', 'hidden', 'year', 'tahun',
            'januari', 'februari', 'maret', 'april', 'mei', 'juni',
            'juli', 'agustus', 'september', 'oktober', 'november', 'desember',
            'january', 'february', 'march', 'april', 'may', 'june',
            'july', 'august', 'september', 'october', 'november', 'december',
            'welcome 2026', 'welcome'
        ]
        
        ws_titles = [ws.title for ws in sh.worksheets()]
        with open('inspect_results.txt', 'w', encoding='utf-8') as f:
            f.write(f"ALL: {ws_titles}\n")
            f.write(f"LOADED: {[t for t in ws_titles if t.lower() not in skip_sheets]}\n")
            f.write(f"SKIPPED: {[t for t in ws_titles if t.lower() in skip_sheets]}\n")
        
    except Exception as e:
        with open('inspect_results.txt', 'w', encoding='utf-8') as f:
            f.write(f"Error: {e}")

if __name__ == "__main__":
    inspect()
