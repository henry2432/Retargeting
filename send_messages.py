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
    print(f"Error initializing Google Sheets: {str(e)}")
    exit(1)

# Wati API 設置
api_token = os.getenv('WATI_API_TOKEN')
base_url = 'https://live-mt-server.wati.io/2601/api/v1'
add_contact_url = f'{base_url}/addContact'
send_message_url = f'{base_url}/sendTemplateMessage'
get_contacts_url = f'{base_url}/getContacts'

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
def check_contact_exists(phone):
    try:
        response = session.get(get_contacts_url, headers=headers, params={'phone': phone}, timeout=15)
        response.raise_for_status()
        return bool(response.json().get('contacts'))
    except requests.RequestException as e:
        print(f"Error checking contact {phone}: {str(e)}")
        return False

# 步驟 1：從 Contacts 分頁加入或更新聯繫人
try:
    contacts_worksheet = sheet.worksheet('Contacts')
    contacts_data = contacts_worksheet.get_all_records()
    if not contacts_data:
        print("No data found in Contacts worksheet")
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
        
        try:
            response = session.post(add_contact_url, json=contact_payload, headers=headers, timeout=15)
            response.raise_for_status()
            if response.status_code in [200, 409]:
                if not row['Added']:
                    contacts_worksheet.update_cell(index + 2, contacts_df.columns.get_loc('Added') + 1, True)
                print(f"Contact {phone} added/updated with AllowBroadcast={allow_broadcast}")
            else:
                print(f"Failed to update contact {phone}: {response.status_code} {response.text}")
        except requests.RequestException as e:
            print(f"Error updating contact {phone}: {str(e)}")
        time.sleep(2)  # 遵守 10 次/10 秒限制
except Exception as e:
    print(f"Error accessing Contacts worksheet: {str(e)}")
    exit(1)

# 步驟 2：從 Rental 和 VIP 分頁傳送訊息
for sheet_name in ['Rental', 'VIP']:
    try:
        worksheet = sheet.worksheet(sheet_name)
        data = worksheet.get_all_records()
        if not data:
            print(f"No data found in {sheet_name} worksheet")
            continue
        df = pd.DataFrame(data)

        df['Sent'] = df['Sent'].astype(bool)
        df['Last_Sent_Date'] = df['Last_Sent_Date'].replace('', None)

        for index, row in df.iterrows():
            if row['Sent']:
                continue
            phone = format_phone(row['Phone'])
            name = row['Name']
            
            if not check_contact_exists(phone):
                print(f"Contact {phone} not found in Wati, skipping message")
                continue
                
            message_payload = {
                'phone': phone,
                'template_name': 'woocommerce_default_follow_up_v2',
                'parameters': {
                    'name': name,
                    'shop_name': 'Kayarine Store'
                }
            }
            try:
                response = session.post(send_message_url, json=message_payload, headers=headers, timeout=15)
                response.raise_for_status()
                if response.status_code == 200:
                    worksheet.update_cell(index + 2, df.columns.get_loc('Sent') + 1, True)
                    worksheet.update_cell(index + 2, df.columns.get_loc('Last_Sent_Date') + 1, 
                                         datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d %H:%M:%S'))
                    print(f"Message sent to {phone} from {sheet_name}")
                else:
                    print(f"Failed to send to {phone} from {sheet_name}: {response.status_code} {response.text}")
            except requests.RequestException as e:
                print(f"Error sending to {phone} from {sheet_name}: {str(e)}")
            time.sleep(0.5)  # 遵守 30 次/10 秒限制
    except Exception as e:
        print(f"Error accessing {sheet_name} worksheet: {str(e)}")
