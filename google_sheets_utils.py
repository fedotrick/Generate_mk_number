import os
import pickle
import json
import datetime
from collections import Counter
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

# ID таблицы Google Sheets (будет установлен при первом запуске)
SPREADSHEET_ID = None

# Имя листа
SHEET_NAME = "Маршрутные карты"

# Области доступа для API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def authenticate_google_sheets():
    """Аутентификация в Google Sheets API с использованием сервисного аккаунта."""
    creds = None
    # Файл credentials.json содержит ключ сервисного аккаунта.
    if os.path.exists('credentials.json'):
        creds = service_account.Credentials.from_service_account_file(
            'credentials.json', scopes=SCOPES)
    else:
        raise FileNotFoundError("Файл credentials.json не найден. Следуйте инструкции по настройке.")
    return creds

def initialize_spreadsheet():
    """Инициализация таблицы Google Sheets"""
    global SPREADSHEET_ID
    
    # Если ID таблицы уже установлен, возвращаем его
    if SPREADSHEET_ID:
        return SPREADSHEET_ID
    
    creds = authenticate_google_sheets()
    service = build('sheets', 'v4', credentials=creds)
    
    try:
        # Проверяем, существует ли файл с ID таблицы
        if os.path.exists('spreadsheet_id.txt'):
            with open('spreadsheet_id.txt', 'r') as f:
                SPREADSHEET_ID = f.read().strip()
                return SPREADSHEET_ID
        
        # Если таблица не существует, создаем новую
        spreadsheet = {
            'properties': {
                'title': 'Маршрутные карты'
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                    fields='spreadsheetId').execute()
        SPREADSHEET_ID = spreadsheet.get('spreadsheetId')
        
        # Сохраняем ID таблицы в файл
        with open('spreadsheet_id.txt', 'w') as f:
            f.write(SPREADSHEET_ID)
        
        # Создаем заголовки
        create_headers(service, SPREADSHEET_ID)
        
        return SPREADSHEET_ID
    except HttpError as err:
        print(f"Ошибка при инициализации таблицы: {err}")
        return None

def create_headers(service, spreadsheet_id):
    """Создание заголовков в таблице"""
    values = [['id', 'Номер_бланка', 'Учетный_номер', 'Номер_кластера', 'Статус', 'Дата_создания', 'Путь_к_файлу']]
    body = {
        'values': values
    }
    try:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=f"'{SHEET_NAME}'!A1",
            valueInputOption="RAW", body=body).execute()
    except HttpError as err:
        print(f"Ошибка при создании заголовков: {err}")

def save_to_google_sheets(form_number, file_path):
    """Сохранение информации о созданной маршрутной карте в Google Sheets"""
    creds = authenticate_google_sheets()
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = initialize_spreadsheet()
    
    if not spreadsheet_id:
        raise Exception("Не удалось инициализировать таблицу Google Sheets")
    
    try:
        # Получаем все данные из таблицы
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"'{SHEET_NAME}'!A:A").execute()
        rows = result.get('values', [])
        
        # Определяем следующий ID
        max_id = 0
        if rows:
            # Пропускаем заголовок
            for row in rows[1:]:
                if row and row[0]:
                    try:
                        max_id = max(max_id, int(row[0]))
                    except ValueError:
                        pass
        
        next_id = max_id + 1
        date_created = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Добавляем новую строку с данными
        values = [[next_id, form_number, "", "", "", date_created, file_path]]
        body = {
            'values': values
        }
        
        # Находим первую пустую строку
        range_name = f"'{SHEET_NAME}'!A{len(rows)+1}:G{len(rows)+1}"
        
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption="RAW", body=body).execute()
            
    except HttpError as err:
        print(f"Ошибка при сохранении в Google Sheets: {err}")
        raise

def check_duplicate_form_number(form_number):
    """Проверка существования бланка с таким номером в Google Sheets"""
    creds = authenticate_google_sheets()
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = initialize_spreadsheet()
    
    if not spreadsheet_id:
        return False
    
    try:
        # Получаем все данные из таблицы
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"'{SHEET_NAME}'!A:G").execute()
        rows = result.get('values', [])
        
        # Если нет данных или только заголовки, дубликатов нет
        if len(rows) <= 1:
            return False
        
        # Проверяем каждую строку, начиная со второй (первая - заголовки)
        for row in rows[1:]:
            if len(row) > 1 and row[1] == form_number:
                return True
        
        return False
    except HttpError as err:
        print(f"Ошибка при проверке дубликатов в Google Sheets: {err}")
        return False

def get_next_form_number():
    """Получение следующего доступного номера маршрутной карты"""
    creds = authenticate_google_sheets()
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = initialize_spreadsheet()
    
    if not spreadsheet_id:
        return "000001"
    
    try:
        # Получаем все данные из таблицы
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"'{SHEET_NAME}'!A:B").execute()
        rows = result.get('values', [])
        
        # Если нет данных или только заголовки, возвращаем начальный номер
        if len(rows) <= 1:
            return "000001"
        
        # Получаем все номера бланков
        form_numbers = []
        for row in rows[1:]:
            if len(row) > 1 and row[1]:
                form_numbers.append(row[1])
        
        # Если нет номеров, возвращаем начальный
        if not form_numbers:
            return "000001"
        
        # Находим максимальный номер и добавляем 1
        max_number = 0
        for form_number in form_numbers:
            try:
                num = int(form_number)
                max_number = max(max_number, num)
            except ValueError:
                continue
        
        next_number = max_number + 1
        return f"{next_number:06d}"
    except HttpError as err:
        print(f"Ошибка при получении следующего номера: {err}")
        return "000001"

def get_sheet_statistics():
    """Получение статистики по записям в Google Sheets."""
    creds = authenticate_google_sheets()
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = initialize_spreadsheet()

    if not spreadsheet_id:
        return None

    try:
        # Читаем всю колонку "Статус" (колонка E)
        range_name = f"'{SHEET_NAME}'!E2:E" # E2, чтобы пропустить заголовок
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_name).execute()
        
        statuses = result.get('values', [])
        
        if not statuses:
            return {'total': 0, 'statuses': {}}
            
        # Уплощаем список и считаем статусы
        status_list = [item for sublist in statuses for item in sublist if item]
        total_records = len(statuses)
        
        # Если все статусы пустые, возвращаем только общее количество
        if not status_list:
             return {'total': total_records, 'statuses': {'Не указан': total_records}}

        status_counts = Counter(status_list)
        
        # Добавляем количество пустых статусов
        empty_statuses = total_records - len(status_list)
        if empty_statuses > 0:
            status_counts['Не указан'] = status_counts.get('Не указан', 0) + empty_statuses

        return {'total': total_records, 'statuses': dict(status_counts)}

    except HttpError as err:
        print(f"Ошибка при получении статистики из Google Sheets: {err}")
        return None
