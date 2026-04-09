import gspread
from google.oauth2.service_account import Credentials
import toml
import os
import re
import pandas as pd

def debug():
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
        
        ws_list = ['📊Promosi Statistik 2026', '🖼️Konten Medsos', '📣Press Release']
        
        with open('debug_output.txt', 'w', encoding='utf-8') as f:
            for ws_name in ws_list:
                try:
                    ws = sh.worksheet(ws_name)
                    data = ws.get_all_values()
                    f.write(f"\n--- SHEET: {ws_name} ---\n")
                    f.write(f"Total Rows: {len(data)}\n")
                    if len(data) > 0:
                        f.write(f"Row 1: {data[0]}\n")
                        f.write(f"Row 2: {data[1]}\n")
                    f.write("-" * 30 + "\n")
                except Exception as e:
                    f.write(f"Error loading {ws_name}: {e}\n")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    debug()
