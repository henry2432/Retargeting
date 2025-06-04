import pandas as pd
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import os
import json
import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Google Sheets API 設置
try:
    credentials = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
    scope = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(credentials, scopes=scope)
    client = gspread.authorize(creds)

    # Google Sheets Spreadsheet ID
    spreadsheet_id = '1dDQCQMipQXzKyxKiljrQjLe6NOwcFOTmGdj-1y7lOp0'
    sheet = client.open_by_key(spreadsheet_id)
except Exception as e:
    print(f"初始化 Google Sheets 失敗: {str(e)}")
    exit(1)

# Wati API 設置
api_token = os.getenv('WATI_API_TOKEN')
base_urls = [
    'https://api.wati.io/api/v1',  # 標準端點
    'https://live-mt-server.wati.io/2601/api/v1'  # 後備端點
]
tenant_id = '2601'

# 配置請求重試與速率限制
session = requests.Session()
retries = Retry(total=3, backoff_factor=2, status_forcelist=[429, 500, 502, 503, 504])
session.mount('https://', HTTPAdapter(max_retries=retries))
headers = {'Authorization': f'Bearer {api_token}', 'Content-Type': 'application/json'}

# 處理電話號碼格式
def format_phone(phone):
    phone = str(phone).strip()
    if not phone.startswith('+'):
        return f'+{phone}'
    return phone

# 檢查聯繫人是否存在
def check_contact_exists(phone, base_url):
    try:
        response = session.get(f'{base_url}/getContacts', headers=headers, params={'phone': phone}, timeout=15)
        response.raise_for_status()
        return bool(response.json().get('contacts'))
    except requests.RequestException as e:
        print(f"檢查聯繫人 {phone} 失敗: {str(e)}")
        return False

# 嘗試不同端點執行請求
def try_request(method, endpoint, base_urls, **kwargs):
    for base_url in base_urls:
        url = f'{base_url}/{endpoint}'
        try:
            response = session.request(method, url, headers=headers, timeout=15, **kwargs)
            response.raise_for_status()
            return response, base_url
        except requests.RequestException as e:
            print(f"請求 {url} 失敗: {response.status_code if 'response' in locals() else ''} {str(e)}")
    return None, None

# 步驟 1：從 Contacts 分頁加入或更新聯繫人
try:
    contacts_worksheet = sheet.worksheet('Contacts')
    contacts_data = contacts_worksheet.get_all_records()
    if not contacts_data:
        print("Contacts 分頁無數據")
        exit(1)
    contacts_df = pd.DataFrame(contacts_data)

    contacts_df['Added'] = contacts_df['Added'].astype(bool)
    contacts_df['AllowBroadcast'] = contacts_df['AllowBroadcast'].astype(bool)

    for index, row in contacts_df.iterrows():
        phone = format_phone(row['Phone'])
        name = row['Name']
        allow_broadcast = row['AllowBroadcast']
        
        contact_payload = {
            'name': name,
            'phone': phone,
            'allowBroadcast': allow_broadcast
        }
        
        response, used_url = try_request('POST', 'addContact', base_urls, json=contact_payload)
        if response and response.status_code in [200, 409]:
            if not row['Added']:
                contacts_worksheet.update_cell(index + 2, contacts_df.columns.get_loc('Added') + 1, True)
            print(f"聯繫人 {phone} 已添加/更新，AllowBroadcast={allow_broadcast}，使用端點: {used_url}")
        else:
            print(f"無法更新聯繫人 {phone}，所有端點均失敗")
        time.sleep(2)  # 遵守 10 次/10 秒限制
except Exception as e:
    print(f"存取 Contacts 分頁失敗: {str(e)}")
    exit(1)

# 步驟 2：從 Rental 和 VIP 分頁傳送訊息
for sheet_name in ['Rental', 'VIP']:
    try:
        worksheet = sheet.worksheet(sheet_name)
        data = worksheet.get_all_records()
        if not data:
            print(f"{sheet_name} 分頁無數據")
            continue
        df = pd.DataFrame(data)

        df['Sent'] = df['Sent'].astype(bool)
        df['Last_Sent_Date'] = df['Last_Sent_Date'].replace('', None)

        for index, row in df.iterrows():
            if row['Sent']:
                continue
            phone = format_phone(row['Phone'])
            name = row['Name']
            
            # 使用成功添加聯繫人的端點檢查
            base_url = base_urls[0]  # 默認標準端點
            if not check_contact_exists(phone, base_url):
                print(f"聯繫人 {phone} 在 Wati 不存在，跳過訊息")
                continue
                
            message_payload = {
                'phone': phone,
                'template_name': 'woocommerce_default_follow_up_v2',
                'parameters': {
                    'name': name,
                    'shop_name': 'Kayarine Store'
                }
            }
            response, used_url = try_request('POST', 'sendTemplateMessage', base_urls, json=message_payload)
            if response and response.status_code == 200:
                worksheet.update_cell(index + 2, df.columns.get_loc('Sent') + 1, True)
                worksheet.update_cell(index + 2, df.columns.get_loc('Last_Sent_Date') + 1, 
                                     datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d %H:%M:%S'))
                print(f"訊息已發送至 {phone}，來自 {sheet_name}，使用端點: {used_url}")
            else:
                print(f"無法發送訊息至 {phone}，來自 {sheet_name}，所有端點均失敗")
            time.sleep(0.5)  # 遵守 30 次/10 秒限制
    except Exception as e:
        print(f"存取 {sheet_name} 分頁失敗: {str(e)}")
