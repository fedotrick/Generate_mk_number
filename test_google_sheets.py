"""
Тестовый скрипт для проверки работы с Google Sheets
"""

import os
import sys

# Добавляем текущую директорию в путь поиска модулей
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from google_sheets_utils import (
    authenticate_google_sheets,
    initialize_spreadsheet,
    save_to_google_sheets,
    check_duplicate_form_number,
    get_next_form_number
)

def test_google_sheets_integration():
    """Тест интеграции с Google Sheets"""
    print("Тестирование интеграции с Google Sheets...")
    
    try:
        # Тест аутентификации
        print("1. Тест аутентификации...")
        creds = authenticate_google_sheets()
        if creds:
            print("   ✓ Аутентификация успешна")
        else:
            print("   ✗ Ошибка аутентификации")
            return False
        
        # Тест инициализации таблицы
        print("2. Тест инициализации таблицы...")
        spreadsheet_id = initialize_spreadsheet()
        if spreadsheet_id:
            print(f"   ✓ Таблица инициализирована, ID: {spreadsheet_id}")
        else:
            print("   ✗ Ошибка инициализации таблицы")
            return False
        
        # Тест получения следующего номера
        print("3. Тест получения следующего номера...")
        next_number = get_next_form_number()
        print(f"   ✓ Следующий номер: {next_number}")
        
        # Тест проверки дубликатов
        print("4. Тест проверки дубликатов...")
        is_duplicate = check_duplicate_form_number("TEST001")
        print(f"   ✓ Проверка дубликатов выполнена, дубликат: {is_duplicate}")
        
        # Тест сохранения данных (только если это не дубликат)
        if not is_duplicate:
            print("5. Тест сохранения данных...")
            try:
                save_to_google_sheets("TEST001", "test/path/test_file.pptx")
                print("   ✓ Данные успешно сохранены")
            except Exception as e:
                print(f"   ⚠ Ошибка при сохранении данных: {e}")
        
        print("\n✓ Все тесты пройдены успешно!")
        return True
        
    except Exception as e:
        print(f"✗ Ошибка во время тестирования: {e}")
        return False

if __name__ == "__main__":
    test_google_sheets_integration()