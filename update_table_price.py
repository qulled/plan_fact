from logging.handlers import RotatingFileHandler

from dotenv import load_dotenv
from googleapiclient import discovery
from google.oauth2 import service_account
from googleapiclient.discovery import build
import logging
import os
import datetime as dt

from api_wb import get_detail_info

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, 'logs/')
log_file = os.path.join(BASE_DIR, 'logs/pars_stocks_table.log')
console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=100000,
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s, %(levelname)s, %(message)s',
    handlers=(
        file_handler,
        console_handler
    )
)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = 'credentials_service.json'
credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE)
service = discovery.build('sheets', 'v4', credentials=credentials)
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')

if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
load_dotenv('.env ')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
NAME_SHEET = os.getenv('NAME_SHEET')


def convert_to_column_letter(column_number):
    column_letter = ''
    while column_number != 0:
        c = ((column_number - 1) % 26)
        column_letter = chr(c + 65) + column_letter
        column_number = (column_number - c) // 26
    return column_letter


def update_table_price(table_id):
    day = date_from.strftime('%d')
    month = date_from.strftime('%m')
    year = date_from.strftime('%Y')
    date = f'{day}.{month}.{year}'
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    result = result.get('values', [])
    row = 2
    body_data = []
    if not result:
        logging.info('No data found.')
    else:
        for values in result:
            if date in values:
                column = values.index(date)+1
            for item in values:
                if item.startswith('Артикул'):
                    article = values[0][-8::]
                    value = get_detail_info(article)['price']
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(column)}{row}',
                         'values': [[f'{value}']]}]
                    row += 13
                    body = {
                        'valueInputOption': 'USER_ENTERED',
                        'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


if __name__ == '__main__':
    date_from = dt.datetime.date(dt.datetime.now())
    range_name = NAME_SHEET
    update_table_price(SPREADSHEET_ID)
