import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import get_as_dataframe
from datetime import datetime
from upload_to_drive import upload_to_drive
import sys
import os
import json

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def load_config(config_path='config.json'):
    path = resource_path(config_path)
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def generate_report(credentials_path = resource_path('credentials.json'),
                    config_path=resource_path('sheets_config.json')):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
        client = gspread.authorize(creds)

        with open(config_path, 'r', encoding='utf-8') as f:
            sheet_configs = json.load(f)

        all_data = []

        for config in sheet_configs:
            sheet = client.open_by_key(config['sheet_id'])
            worksheet = sheet.get_worksheet(0)
            df = get_as_dataframe(worksheet, evaluate_formulas=True)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            df['部門'] = config['department']
            all_data.append(df)

        if not all_data:
            return False, "沒有任何資料可合併"

        final_df = pd.concat(all_data, ignore_index=True)
        today = datetime.today().strftime('%Y%m%d')

        home_dir = os.path.expanduser('~')
        output_dir = os.path.join(home_dir, 'AutoSheetReport_output')
        os.makedirs(output_dir, exist_ok=True)

        base_name = f'部門彙總報表_{today}'
        filename = get_next_filename(base_name, output_dir)

        final_df.to_excel(filename, index=False)

        success, upload_msg = upload_to_drive(filename)
        print(upload_msg)
        return True, f"報表已產出：{filename}\n{upload_msg}"

    except Exception as e:
        return False, f"報表產生失敗：{str(e)}"



def get_next_filename(base_name, output_dir):
    version = 1
    while True:
        versioned_name = f"{base_name}_v{version}.xlsx"
        full_path = os.path.join(output_dir, versioned_name)
        if not os.path.exists(full_path):
            return full_path
        version += 1
