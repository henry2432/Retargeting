import pandas as pd
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timezone, timedelta
import os
import json

# Google Sheets API 設置
credentials = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials, scope)
client = gspread.authorize(creds)

# Google Sheets 名稱
sheet_name = os.getenv('SHEET_NAME', 'Whatsapp Marketing for Walk-in')
sheet = client.open(sheet_name)

# Wati API 設置
api_token = os.getenv('WATI_API_TOKEN')
wati_url = 'https://live-mt-server.wati.io/2601/api/v1/sendTemplateMessage'

# 模板映射
template_map = {
    'Rental': ['rental_promo_en', 'rental_promo_zh'],
    'VIP': ['vip_store_promo_en', 'vip_store_promo_zh']
}

# 處理每個分頁
for sheet_name in ['Rental', 'VIP']:
    worksheet = sheet.worksheet(sheet_name)
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)

    # 確保欄位正確
    df['Sent'] = df['Sent'].astype(bool)
    df['Last_Sent_Date'] = df['Last_Sent_Date'].replace('', None)

    for index, row in df.iterrows():
        if row['Sent']:
            continue
        template_name = template_map[sheet_name][0 if row['Language'] == 'en' else 1]
        payload = {
            'phone': str(row['Phone']),
            'template_name': template_name,
            'parameters': {
                'name': row['Name'],
                'coupon_code': row['Coupon_Code'],
                'custom_message': row.get('Custom_Message', '')
            }
        }
        headers = {'Authorization': f'Bearer {api_token}'}
        response = requests.post(wati_url, json=payload, headers=headers)
        if response.status_code == 200:
            worksheet.update_cell(index + 2, df.columns.get_loc('Sent') + 1, True)
            worksheet.update_cell(index + 2, df.columns.get_loc('Last_Sent_Date') + 1, 
                                 datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d %H:%M:%S'))
        else:
            print(f"Failed to send to {row['Phone']}: {response.text}")
