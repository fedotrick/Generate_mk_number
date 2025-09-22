
import os
from googleapiclient.discovery import build
from google.oauth2 import service_account

# --- The same authentication from the project ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'] # Read-only is enough
creds = None
if os.path.exists('credentials.json'):
    creds = service_account.Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
else:
    print("!!! Ошибка: файл credentials.json не найден.")
    exit()

# --- Read the spreadsheet ID from the file ---
spreadsheet_id = None
if os.path.exists('spreadsheet_id.txt'):
    with open('spreadsheet_id.txt', 'r') as f:
        spreadsheet_id = f.read().strip()
else:
    print("!!! Ошибка: файл spreadsheet_id.txt не найден.")
    exit()

# --- Connect to the API and get spreadsheet properties ---
try:
    service = build('sheets', 'v4', credentials=creds)
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    
    print("--- DEBUG INFO ---")
    print(f"Подключился к таблице с ID: {spreadsheet_id}")
    print("Найдены следующие листы (вкладки):")
    
    if not sheets:
        print("- Внутри таблицы не найдено ни одного листа.")
    else:
        for sheet in sheets:
            sheet_title = sheet.get("properties", {}).get("title", "ИМЯ НЕ НАЙДЕНО")
            print(f"- '{sheet_title}'")
    
    print("------------------")
    print("\nСравните имя листа, которое вы видите выше, с тем, что требует программа: 'Маршрутные карты'")

except Exception as e:
    print(f"\n!!! Произошла ошибка при попытке получить информацию о листах: {e}")
