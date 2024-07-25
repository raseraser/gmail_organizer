import os
import json
import re
import time
import argparse
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
from datetime import datetime, timedelta, timezone
import PyPDF2
import binascii
import quopri
import io

VERSION = '1.2'
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
    results = service.users().messages().list(userId='me', labelIds=[label], maxResults=500).execute()
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

def get_message_datestr(message):
    """
    從 Gmail 訊息中提取日期並將其格式化為指定的字串格式。

    :param message: Gmail 訊息對象
    :param dateformat: 日期字串格式，預設為 'yyyymmdd'
    :return: 格式化的日期字串，或 None 如果沒有找到日期
    """
    headers = message['payload']['headers']
    date_header = next((header for header in headers if header['name'] == 'Date'), None)
    if date_header:
        date_str = date_header['value']
        # 移除可能存在的額外時區信息
        date_str = re.sub(r'\s*\([A-Z]+\)$', '', date_str)
        try:
            # 嘗試解析標準格式
            date_obj = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z')
        except ValueError:
            try:
                # 如果失敗，嘗試解析沒有時區信息的格式
                date_obj = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S')
                # 假設是本地時間，將其轉換為 UTC
                date_obj = date_obj.replace(tzinfo=timezone.utc)
            except ValueError:
                # 如果還是失敗，返回 None 或者原始字符串
                return date_str
        return date_obj.strftime('%Y%m%d_%H%M%S')
    return None

    
def sanitize_filename(filename):
    """
    移除檔名中的非法字元
    """
    return re.sub(r'[\/:*?"<>|]', '', filename)

def save_mail_as_eml(service, msg_id, save_dir, mail_datestr):
    try:
        message = service.users().messages().get(userId='me', id=msg_id, format='raw').execute()
        msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
        mime_msg = email.message_from_bytes(msg_str)

        # 使用日期字符串作為文件名前缀
        filename = f"{mail_datestr}_email.eml"
        filepath = os.path.join(save_dir, filename)

        with open(filepath, 'wb') as f:
            f.write(msg_str)
        print(f"已保存郵件: {filepath}")
    except Exception as error:
        print(f'保存郵件為 .eml 文件時發生錯誤: {error}')

def save_mail_attachments(service, msg_id, save_dir, file_types=None, save_mail=False, save_attachment=True):
    try:
        message = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
        mail_datestr = get_message_datestr(message) or ""

        print(f"處理郵件 ID: {msg_id}")
        print(f"郵件日期: {mail_datestr}")

        if save_mail:
            save_mail_as_eml(service, msg_id, save_dir, mail_datestr)

        if save_attachment:
            process_parts(service, msg_id, message['payload'], save_dir, mail_datestr, file_types)

    except Exception as error:
        print(f'處理郵件附件時發生錯誤: {error}')

def process_parts(service, msg_id, part, save_dir, mail_datestr, file_types, level=0):
    print("  " * level + f"處理部分 MIME 類型: {part.get('mimeType', 'Unknown')}")

    if 'parts' in part:
        for sub_part in part['parts']:
            process_parts(service, msg_id, sub_part, save_dir, mail_datestr, file_types, level + 1)
    elif part.get('mimeType') == 'application/octet-stream':
        print("  " * level + "檢測到 application/octet-stream，調用 handle_octet_stream...")
        handle_octet_stream(service, msg_id, part, save_dir, mail_datestr, file_types)
    elif 'filename' in part.get('body', {}):
        filename = part['body']['filename']
        print("  " * level + f"檢測到文件名: {filename}")
        # 處理其他類型的附件...
    else:
        print("  " * level + "未知部分類型，跳過")

def handle_octet_stream(service, msg_id, part, save_dir, mail_datestr, file_types):
    print("開始處理 octet-stream...")
    if 'data' in part['body']:
        data = part['body']['data']
    else:
        att_id = part['body']['attachmentId']
        print(f"使用 attachmentId: {att_id}")
        att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
        data = att['data']
    
    file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
    
    print(f"數據大小: {len(file_data)} 字節")
    print(f"前100個字節: {binascii.hexlify(file_data[:100])}")
    
    # 保存原始數據以供進一步分析
    raw_filename = f"{mail_datestr}_raw_data.bin"
    raw_filepath = os.path.join(save_dir, raw_filename)
    with open(raw_filepath, 'wb') as f:
        f.write(file_data)
    print(f"已保存原始數據: {raw_filepath}")

    # 檢查是否為 PDF 文件
    if file_data[:4] == b'%PDF':
        print("檢測到 PDF 文件")
        pdf_filename = f"{mail_datestr}_document.pdf"
        pdf_filepath = os.path.join(save_dir, pdf_filename)
        with open(pdf_filepath, 'wb') as f:
            f.write(file_data)
        print(f"已保存 PDF 文件: {pdf_filepath}")
        return

    # 如果不是 PDF，嘗試其他方法
    try:
        print("嘗試作為 email 附件解析...")
        decoded_content = quopri.decodestring(file_data)
        msg = email.message_from_bytes(decoded_content)
        
        for attachment in msg.walk():
            print(f"附件類型: {attachment.get_content_type()}")
            if attachment.get_content_maintype() == 'application' and attachment.get_filename():
                filename = attachment.get_filename()
                print(f"在 octet-stream 中找到附件: {filename}")
                
                file_extension = os.path.splitext(filename)[1].lower()
                if file_types is None or file_extension in file_types:
                    attachment_data = attachment.get_payload(decode=True)
                    save_file(attachment_data, filename, save_dir, mail_datestr)
                else:
                    print(f"跳過不符合類型的附件: {filename}")
    except Exception as e:
        print(f"以 email 附件方式解析失敗: {str(e)}")
    
    # 如果上述方法都失敗，保存為未知二進制文件
    print("保存為未知二進制文件...")
    unknown_filename = f"{mail_datestr}_unknown_content.bin"
    unknown_filepath = os.path.join(save_dir, unknown_filename)
    with open(unknown_filepath, 'wb') as f:
        f.write(file_data)
    print(f"已保存未知內容: {unknown_filepath}")

def save_file(data, filename, save_dir, mail_datestr):
    filepath = os.path.join(save_dir, filename)
    with open(filepath, 'wb') as f:
        f.write(data)
    print(f"已保存文件: {filepath}")

def print_message_structure(part, level=0):
    print('  ' * level + f"MIME Type: {part.get('mimeType', 'Unknown')}")
    if 'filename' in part.get('body', {}):
        print('  ' * level + f"Filename: {part['body'].get('filename', 'Unknown')}")
    if 'parts' in part:
        for sub_part in part['parts']:
            print_message_structure(sub_part, level + 1)
    # 添加更多細節
    if part.get('mimeType') == 'application/octet-stream':
        print('  ' * (level+1) + "Octet-stream details:")
        print('  ' * (level+2) + f"Size: {part['body'].get('size', 'Unknown')} bytes")
        print('  ' * (level+2) + f"AttachmentId: {part['body'].get('attachmentId', 'Not present')}")

def save_attachment(service, msg_id, part, save_dir, mail_datestr, file_types):
    filename = part['body']['filename']
    print(f"正在處理附件: {filename} ...")
    file_extension = os.path.splitext(filename)[1].lower()
    
    if file_types is None or file_extension in file_types:
        if 'data' in part['body']:
            data = part['body']['data']
        else:
            att_id = part['body']['attachmentId']
            att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
            data = att['data']
        file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
        
        full_filename = f"{mail_datestr}_{filename}"
        filepath = os.path.join(save_dir, full_filename)
        
        with open(filepath, 'wb') as f:
            f.write(file_data)
        print(f"已保存附件: {filepath}")
    else:
        print(f"跳過不符合類型的附件: {filename}")

def delete_message(service, message_id):
    service.users().messages().trash(userId='me', id=message_id).execute()

def load_config(cfg_file='config.json'):
    if not os.path.exists(cfg_file):
        config = {
            "label": {
                "value": "需要處理",
                "comment": "要處理的郵件標籤"
            },
            "confirm_each_run": {
                "value": True,
                "comment": "是否在每次運行時要求確認"
            },
            "download_mail_attachments": {
                "enabled": {
                    "value": True,
                    "comment": "是否啟用"
                },
                "save_mail": {
                    "value": False,
                    "comment": "是否保存信件檔案(.eml)"
                },
                "save_attachment": {
                    "value": True,
                    "comment": "是否下載附件"
                },
                "file_types": {
                    "value": [".pdf", ".xlsx", ".docx"],
                    "comment": "要下載的文件類型列表"
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
        with open(cfg_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print(f"已創建默認配置文件 {cfg_file}，請根據需要修改它。")
    
    with open(cfg_file, 'r', encoding='utf-8') as f:
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
    results = service.users().messages().list(userId='me', labelIds=[label_id], maxResults=500).execute()
    return results.get('messages', [])

def display_progress(current, total, start_time):
    elapsed_time = time.time() - start_time
    progress = current / total
    eta = elapsed_time / progress - elapsed_time if progress > 0 else 0
    
    bar_length = 20
    filled_length = int(bar_length * progress)
    bar = '=' * filled_length + '-' * (bar_length - filled_length)
    
    print(f'\r處理進度: [{bar}] {current}/{total} 封郵件 - {progress:.1%} 完成 - 預計剩餘時間: {eta:.0f}秒', end='', flush=True)

def main(cfg_file):
    config = load_config(cfg_file)
    service = get_gmail_service()

    label_name = config['label']['value']
    label_id = get_label_id(service, label_name)
    if label_id is None:
        print(f"錯誤：找不到名為 '{label_name}' 的標籤。")
        return

    messages = get_messages_with_label(service, label_id)
    total_messages = len(messages)

    print(f"找到 {total_messages} 封帶有標籤 '{label_name}' 的郵件。 (gmail系統最多只能取得500封郵件，如果有更多郵件，請多次運行本程序。)")
    print(f"數量多時，可能需要較長時間處理 ...")

    download_config = config['download_mail_attachments']
    delete_config = config['delete_emails']

    download_count = 0
    delete_count = 0

    if download_config['enabled']['value']:
        download_start = parse_date(download_config['date_range']['from']['value'])
        download_end = parse_date(download_config['date_range']['to']['value'])
        save_mail = download_config['save_mail']['value']
        save_attachment = download_config['save_attachment']['value']
        save_dir = download_config['save_directory']['value']
        # 如果輸出目錄不存在，則創建它
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

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
        print(f"\n處理郵件: {message['id']}")
        print(f"郵件摘要: {message.get('snippet', 'No snippet available')[:100]}...")

        message_date = get_message_date(service, message['id'])
        if message_date is None:
            continue

        if download_config['enabled']['value'] and is_within_date_range(message_date, download_start, download_end):
            save_mail_attachments(
                service, 
                message['id'], 
                save_dir, 
                file_types=download_config['file_types']['value'] if 'file_types' in download_config else None,
                save_mail=download_config['save_mail']['value'],
                save_attachment=download_config['save_attachment']['value']
            )
        if delete_config['enabled']['value'] and is_within_date_range(message_date, delete_start, delete_end):
            delete_message(service, message['id'])
        
        processed_count += 1
        display_progress(processed_count, total_messages, start_time)

    print("\n操作完成。")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('cfg_file', type=str, help='Path to the configuration file')

    args = parser.parse_args()
    main(args.cfg_file)