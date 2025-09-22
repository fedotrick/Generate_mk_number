
import sqlite3
import os

DB_FILE = "маршрутные_карты.db"

if not os.path.exists(DB_FILE):
    print(f"Файл базы данных '{DB_FILE}' не найден!")
    exit()

try:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # Get list of tables
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()

    if not tables:
        print("В базе данных не найдено таблиц.")
    else:
        print(f"Найдены таблицы в '{DB_FILE}':")
        for table_name_tuple in tables:
            table_name = table_name_tuple[0]
            print(f"\n--- Схема для таблицы '{table_name}' ---")
            
            # Get table schema
            cursor.execute(f"PRAGMA table_info({table_name});")
            columns = cursor.fetchall()
            
            if not columns:
                print("Не удалось получить информацию о столбцах.")
            else:
                for column in columns:
                    # column format: (id, name, type, notnull, default_value, pk)
                    print(f"  - {column[1]} ({column[2]})")

except sqlite3.Error as e:
    print(f"Произошла ошибка SQLite: {e}")
finally:
    if conn:
        conn.close()

