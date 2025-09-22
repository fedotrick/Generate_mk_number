@echo off
title Тест Google Sheets
echo Запуск теста интеграции с Google Sheets...
echo.
echo Перед запуском убедитесь, что:
echo 1. Файл credentials.json находится в корневой папке проекта
echo 2. Установлены все зависимости из requirements.txt
echo.
pause
echo.
python test_google_sheets.py
echo.
pause