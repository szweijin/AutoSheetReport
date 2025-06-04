import os
import json
import re
import sys
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # PyInstaller 解壓臨時目錄
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_config(config_path='config.json'):
    path = resource_path(config_path)
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def get_next_drive_filename(service, folder_id, base_name):
    query = f"'{folder_id}' in parents and trashed = false"
    files = []
    page_token = None
    while True:
        results = service.files().list(
            q=query,
            fields="nextPageToken, files(name)",
            pageToken=page_token
        ).execute()
        files.extend(results.get('files', []))
        page_token = results.get('nextPageToken', None)
        if page_token is None:
            break

    pattern = re.compile(rf"{re.escape(base_name)}_v(\d+)\.xlsx")
    max_version = 0
    for file in files:
        match = pattern.match(file['name'])
        if match:
            version = int(match.group(1))
            if version > max_version:
                max_version = version

    next_version = max_version + 1
    return f"{base_name}_v{next_version}.xlsx"

def upload_to_drive(file_path, credentials_path='credentials.json', config_path='config.json'):
    try:
        config = load_config(config_path)
        folder_id = config['drive_folder_id']

        SCOPES = ['https://www.googleapis.com/auth/drive']
        creds_path = resource_path(credentials_path)
        creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        service = build('drive', 'v3', credentials=creds)

        base_name = os.path.splitext(os.path.basename(file_path))[0].rsplit('_v', 1)[0]
        upload_name = get_next_drive_filename(service, folder_id, base_name)

        file_metadata = {
            'name': upload_name,
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True, f"已上傳至 Google Drive：{upload_name}（檔案 ID：{file.get('id')}）"

    except Exception as e:
        return False, f"上傳失敗：{e}"
