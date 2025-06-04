import pandas as pd
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import os
import json

# Google Sheets API 設置（僅使用 Sheets API）
credentials = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
scope = ['https://www.googleapis.com/auth/spreadsheets']  # 僅 Sheets API
creds = Credentials.from_service_account_info(credentials, scopes=scope)
client = gspread.authorize(creds)

# Google Sheets 名稱
sheet_name = os.getenv('SHEET_NAME', 'Whatsapp Marketing for Walk-in')
sheet = client.open(sheet_name)

# Wati API 設置
api_token = os.getenv('WATI_API_TOKEN')
wati_url = 'https://live-mt-server.wati.io/2601/api/v1/sendTemplateMessage'

# 處理 Rental 分頁（測試）
worksheet = sheet.worksheet('Rental')
data = worksheet.get_all_records()
df = pd.DataFrame(data)

# 確保欄位正確
df['Sent'] = df['Sent'].astype(bool)
df['Last_Sent_Date'] = df['Last_Sent_Date'].replace('', None)

for index, row in df.iterrows():
    if row['Sent']:
        continue
    payload = {
        'phone': str(row['Phone']),
        'template_name': 'woocommerce_default_follow_up_v2',
        'parameters': {
            'name': row['Name'],
            'shop_name': row['Shop_Name']
        }
    }
    headers = {'Authorization': f'Bearer {api_token}'}
    response = requests.post(wati_url, json=payload, headers=headers)
    if response.status_code == 200:
        worksheet.update_cell(index + 2, df.columns.get_loc('Sent') + 1, True)
        worksheet.update_cell(index + 2, df.columns.get_loc('Last_Sent_Date') + 1, 
                             datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d %H:%M:%S'))
        print(f"Message sent to {row['Phone']}")
    else:
        print(f"Failed to send to {row['Phone']}: {response.text}")
