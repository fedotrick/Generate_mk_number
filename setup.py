import sys
from cx_Freeze import setup, Executable

# Зависимости
build_exe_options = {
    "packages": ["os", "qrcode", "pptx", "openpyxl", "kivy", "datetime"],
    "excludes": [],
    "include_files": [
        ("ШАБЛОН.pptx", "ШАБЛОН.pptx"), 
        ("app.ico", "app.ico")
    ]
}

# Базовый EXE файл (для Windows без терминала)
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Генератор Маршрутных Карт",
    version="1.0",
    description="Приложение для генерации маршрутных карт",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "generate_form_pptx.py", 
            base=base,
            target_name="Генератор Маршрутных Карт.exe",
            icon="app.ico"
        )
    ]
)
