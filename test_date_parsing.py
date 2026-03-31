import pandas as pd
import re

def clean_indo_month_string(date_str):
    if not isinstance(date_str, str): return date_str
    indo_to_eng = {
        'januari': 'January', 'februari': 'February', 'maret': 'March',
        'april': 'April', 'mei': 'May', 'juni': 'June',
        'juli': 'July', 'agustus': 'August', 'september': 'September',
        'oktober': 'October', 'november': 'November', 'desember': 'December'
    }
    
    original_str = date_str.strip()
    if not original_str: return date_str
    
    low_str = original_str.lower()
    for indo, eng in indo_to_eng.items():
        if indo in low_str:
            original_str = re.sub(re.escape(indo), eng, original_str, flags=re.IGNORECASE)
            break
            
    if not re.search(r'\d{4}', original_str):
        original_str = f"{original_str} 2026"
        
    return original_str

test_cases = [
    "15 Mei 2026",
    "2 Agustus",
    "31 Maret 2026",
    "April",
    "Mei",
    "Tayang: 15 Juni",
    "2026-05-15"
]

print("Original -> Cleaned -> Parsed")
for tc in test_cases:
    cleaned = clean_indo_month_string(tc)
    parsed = pd.to_datetime(cleaned, errors='coerce', dayfirst=True)
    print(f"{tc} -> {cleaned} -> {parsed}")
