import os
import json
import re
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
from datetime import datetime, timedelta

VERSION = '1.0'
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

def get_gmail_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return build('gmail', 'v1', credentials=creds)

def get_messages_with_label(service, label):
    results = service.users().messages().list(userId='me', labelIds=[label]).execute()
    return results.get('messages', [])

def get_message_date(service, message_id):
    message = service.users().messages().get(userId='me', id=message_id, format='metadata', metadataHeaders=['Date']).execute()
    date_str = next(header['value'] for header in message['payload']['headers'] if header['name'] == 'Date')
    
    # 清理日期字符串
    date_str = re.sub(r'\s*\([A-Z]+\)$', '', date_str)
    
    try:
        # 嘗試解析帶有時區信息的日期
        return datetime.strptime(date_str, "%a, %d %b %Y %H:%M:%S %z").date()
    except ValueError:
        try:
            # 如果失敗，嘗試解析不帶時區信息的日期
            return datetime.strptime(date_str, "%a, %d %b %Y %H:%M:%S").date()
        except ValueError:
            # 如果還是失敗，打印錯誤信息並返回 None
            print(f"無法解析日期: {date_str}")
            return None

def is_within_date_range(message_date, start_date, end_date):
    return start_date <= message_date <= end_date

def save_attachments(service, message_id, save_dir):
    message = service.users().messages().get(userId='me', id=message_id).execute()
    for part in message['payload'].get('parts', []):
        if part.get('filename'):
            if 'data' in part['body']:
                data = part['body']['data']
            else:
                att_id = part['body']['attachmentId']
                att = service.users().messages().attachments().get(userId='me', messageId=message_id, id=att_id).execute()
                data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            
            file_path = os.path.join(save_dir, part['filename'])
            with open(file_path, 'wb') as f:
                f.write(file_data)

def delete_message(service, message_id):
    service.users().messages().trash(userId='me', id=message_id).execute()

def load_config():
    if not os.path.exists('config.json'):
        config = {
            "label": {
                "value": "需要處理",
                "comment": "要處理的郵件標籤"
            },
            "confirm_each_run": {
                "value": True,
                "comment": "是否在每次運行時要求確認"
            },
            "download_attachments": {
                "enabled": {
                    "value": True,
                    "comment": "是否下載附件"
                },
                "save_directory": {
                    "value": "attachments",
                    "comment": "保存附件的目錄"
                },
                "date_range": {
                    "from": {
                        "value": "2023-01-01",
                        "comment": "下載附件的開始日期 (YYYY-MM-DD)"
                    },
                    "to": {
                        "value": "2023-12-31",
                        "comment": "下載附件的結束日期 (YYYY-MM-DD)"
                    }
                }
            },
            "delete_emails": {
                "enabled": {
                    "value": True,
                    "comment": "是否刪除郵件"
                },
                "date_range": {
                    "from": {
                        "value": "2023-01-01",
                        "comment": "刪除郵件的開始日期 (YYYY-MM-DD)"
                    },
                    "to": {
                        "value": "2023-12-31",
                        "comment": "刪除郵件的結束日期 (YYYY-MM-DD)"
                    }
                }
            }
        }
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print("已創建默認配置文件 config.json，請根據需要修改它。")
    
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def parse_date(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d").date()

def get_label_id(service, label_name):
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    for label in labels:
        if label['name'] == label_name:
            return label['id']
    return None

def get_messages_with_label(service, label_id):
    results = service.users().messages().list(userId='me', labelIds=[label_id]).execute()
    return results.get('messages', [])

def display_progress(current, total, start_time):
    elapsed_time = time.time() - start_time
    progress = current / total
    eta = elapsed_time / progress - elapsed_time if progress > 0 else 0
    
    bar_length = 20
    filled_length = int(bar_length * progress)
    bar = '=' * filled_length + '-' * (bar_length - filled_length)
    
    print(f'\r處理進度: [{bar}] {current}/{total} 封郵件 - {progress:.1%} 完成 - 預計剩餘時間: {eta:.0f}秒', end='', flush=True)

def main():
    config = load_config()
    service = get_gmail_service()

    label_name = config['label']['value']
    label_id = get_label_id(service, label_name)
    if label_id is None:
        print(f"錯誤：找不到名為 '{label_name}' 的標籤。")
        return

    messages = get_messages_with_label(service, label_id)
    total_messages = len(messages)

    print(f"找到 {total_messages} 封帶有標籤 '{label_name}' 的郵件。")

    download_config = config['download_attachments']
    delete_config = config['delete_emails']

    download_count = 0
    delete_count = 0

    if download_config['enabled']['value']:
        download_start = parse_date(download_config['date_range']['from']['value'])
        download_end = parse_date(download_config['date_range']['to']['value'])
        save_dir = download_config['save_directory']['value']

    if delete_config['enabled']['value']:
        delete_start = parse_date(delete_config['date_range']['from']['value'])
        delete_end = parse_date(delete_config['date_range']['to']['value'])

    # 在處理郵件的循環中使用 message['id']
    for message in messages:
        message_date = get_message_date(service, message['id'])
        if message_date is None:
            continue  # 跳過無法解析日期的郵件
        
        if download_config['enabled']['value']:
            if is_within_date_range(message_date, download_start, download_end):
                download_count += 1

        if delete_config['enabled']['value']:
            if is_within_date_range(message_date, delete_start, delete_end):
                delete_count += 1

    if download_config['enabled']['value']:
        print(f"符合「下載附件」條件（{download_start} 到 {download_end}）的郵件有 {download_count} 封")
    if delete_config['enabled']['value']:
        print(f"符合「刪除」條件（{delete_start} 到 {delete_end}）的郵件有 {delete_count} 封")

    if config['confirm_each_run']['value']:
        confirm = input("是否確認執行以上操作？(y/n): ")
        if confirm.lower() != 'y':
            print("操作已取消。")
            return

    start_time = time.time()
    processed_count = 0

    for message in messages:
        message_date = get_message_date(service, message['id'])
        if message_date is None:
            continue

        if download_config['enabled']['value'] and is_within_date_range(message_date, download_start, download_end):
            save_attachments(service, message['id'], save_dir)
        
        if delete_config['enabled']['value'] and is_within_date_range(message_date, delete_start, delete_end):
            delete_message(service, message['id'])
        
        processed_count += 1
        display_progress(processed_count, total_messages, start_time)

    print("\n操作完成。")

if __name__ == '__main__':
    main()