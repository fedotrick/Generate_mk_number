
import sqlite3
import os
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError

# --- Configuration ---
DB_FILE = "маршрутные_карты.db"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_NAME = "Маршрутные карты"

# --- 1. Get Spreadsheet ID ---
spreadsheet_id = None
if os.path.exists('spreadsheet_id.txt'):
    with open('spreadsheet_id.txt', 'r') as f:
        spreadsheet_id = f.read().strip()
else:
    print("!!! Ошибка: файл spreadsheet_id.txt не найден.")
    exit()

# --- 2. Authenticate with Google ---
creds = None
if os.path.exists('credentials.json'):
    creds = service_account.Credentials.from_service_account_file(
        'credentials.json', scopes=SCOPES)
else:
    print("!!! Ошибка: файл credentials.json не найден.")
    exit()

try:
    service = build('sheets', 'v4', credentials=creds)
    print("Успешное подключение к Google Sheets API.")

    # --- 3. Read existing data from Google Sheet ---
    print("Читаю существующие данные из Google Таблицы, чтобы избежать дубликатов...")
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f"'{SHEET_NAME}'!B:B").execute()
    values = result.get('values', [])
    
    existing_form_numbers = set()
    if values:
        # Flatten list of lists and add to set
        for row in values:
            if row:
                existing_form_numbers.add(row[0])
    print(f"Найдено {len(existing_form_numbers)} существующих записей в Google Таблице.")

    # --- 4. Read data from SQLite DB ---
    if not os.path.exists(DB_FILE):
        print(f"!!! Ошибка: Файл базы данных '{DB_FILE}' не найден!")
        exit()

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT id, Номер_бланка, Учетный_номер, Номер_кластера, Статус, Дата_создания, Путь_к_файлу FROM маршрутные_карты")
    db_rows = cursor.fetchall()
    conn.close()
    print(f"Найдено {len(db_rows)} записей в локальной базе данных.")

    # --- 5. Filter data to be migrated ---
    rows_to_migrate = []
    for row in db_rows:
        # Assuming row[1] is 'Номер_бланка'
        if row[1] not in existing_form_numbers:
            rows_to_migrate.append(list(row)) # Convert tuple to list

    print(f"Будет перенесено {len(rows_to_migrate)} новых записей.")

    # --- 6. Append new data to Google Sheet ---
    if not rows_to_migrate:
        print("Миграция не требуется. Все данные уже синхронизированы.")
    else:
        print("Добавляю новые записи в Google Таблицу...")
        body = {
            'values': rows_to_migrate
        }
        # Use append to add rows to the first empty line
        result = sheet.values().append(
            spreadsheetId=spreadsheet_id,
            range=f"'{SHEET_NAME}'!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body).execute()
        
        print("\n--- УСПЕХ! ---")
        updated_range = result.get('updates', {}).get('updatedRange')
        print(f"Данные успешно добавлены в диапазон: {updated_range}")
        print(f"Всего перенесено: {len(rows_to_migrate)} записей.")

except HttpError as err:
    print(f"\n!!! Произошла ошибка API Google: {err}")
except sqlite3.Error as e:
    print(f"\n!!! Произошла ошибка SQLite: {e}")
except Exception as e:
    print(f"\n!!! Произошла непредвиденная ошибка: {e}")
