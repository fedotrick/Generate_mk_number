@echo off
title Генератор маршрутных карт
echo Запуск Streamlit приложения "Генератор маршрутных карт"...
echo.
echo Приложение будет доступно в браузере по адресу: http://localhost:8502
echo Для остановки приложения нажмите Ctrl+C
echo.
streamlit run streamlit_app.py --server.port 8502
pause