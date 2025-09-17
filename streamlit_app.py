import streamlit as st
import os
import sys

# Add the current directory to the path so we can import our modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generate_form_pptx_core import (
    create_database,
    save_to_database,
    check_duplicate_form_number,
    update_empty_to_null,
    generate_form_with_qr,
    generate_multiple_forms,
    get_next_form_number
)

# Initialize the database
create_database()
try:
    update_empty_to_null()
except Exception as e:
    st.warning(f"Предупреждение при обновлении базы данных: {e}")

# Set page config for a more modern look
st.set_page_config(
    page_title="Генератор маршрутных карт",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for a more modern design
st.markdown("""
<style>
    .stApp {
        background-color: #f0f2f6;
    }
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 8px 8px 0 0;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #667eea;
        color: white;
    }
    .generate-button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border: none;
        color: white;
        padding: 15px 32px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 8px;
        transition: all 0.3s ease;
    }
    .generate-button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    .info-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .error-box {
        background-color: #f8d7da;
        border-left: 5px solid #dc3545;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border-left: 5px solid #ffc107;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Main header
st.markdown("<div class='main-header'><h1>📋 Генератор маршрутных карт</h1><p>Создание маршрутных карт в формате PowerPoint с QR-кодами</p></div>", unsafe_allow_html=True)

# Get the next form number to suggest
next_form_number = get_next_form_number()

# Create tabs for single and multiple form generation
tab1, tab2 = st.tabs(["Один бланк", "Несколько бланков"])

with tab1:
    st.header("Создание одной маршрутной карты")
    
    # Information card
    st.markdown("<div class='info-card'><h3>Информация</h3><p>Следующий доступный номер маршрутной карты: <strong>" + next_form_number + "</strong></p></div>", unsafe_allow_html=True)
    
    # Input for form number with suggestion
    form_number = st.text_input("Номер маршрутной карты", value=next_form_number, help="Введите номер маршрутной карты или используйте предложенный")
    
    # Template file
    template_path = st.text_input("Путь к шаблону PowerPoint", value="ШАБЛОН.pptx", help="Укажите путь к файлу шаблона PowerPoint")
    
    # Output directory
    output_dir = st.text_input("Директория для сохранения файлов", value="Маршрутные_карты", help="Укажите директорию для сохранения сгенерированных файлов")
    
    # Generate button with custom styling
    if st.button("Создать маршрутную карту", key="single_generate"):
        if not os.path.exists(template_path):
            st.markdown("<div class='error-box'>❌ Файл шаблона '" + template_path + "' не найден!</div>", unsafe_allow_html=True)
        else:
            try:
                # Create output directory if it doesn't exist
                os.makedirs(output_dir, exist_ok=True)
                
                # Format the form number
                form_number_formatted = f"{int(form_number):06d}"
                output_path = os.path.join(output_dir, f"маршрутная_карта_{form_number_formatted}.pptx")
                
                # Generate the form
                generate_form_with_qr(template_path, output_path, form_number_formatted)
                st.markdown("<div class='success-box'>✅ Маршрутная карта успешно создана: " + output_path + "</div>", unsafe_allow_html=True)
                
                # Provide download link
                with open(output_path, "rb") as file:
                    st.download_button(
                        label="📥 Скачать файл",
                        data=file,
                        file_name=f"маршрутная_карта_{form_number_formatted}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="download_single"
                    )
            except Exception as e:
                st.markdown("<div class='error-box'>❌ Ошибка при создании маршрутной карты: " + str(e) + "</div>", unsafe_allow_html=True)

with tab2:
    st.header("Создание нескольких маршрутных карт")
    
    # Information card
    st.markdown("<div class='info-card'><h3>Информация</h3><p>Следующий доступный номер маршрутной карты: <strong>" + next_form_number + "</strong></p></div>", unsafe_allow_html=True)
    
    # Starting number with suggestion
    start_number = st.text_input("Начальный номер", value=next_form_number, help="Введите начальный номер для генерации нескольких карт")
    
    # Count
    count = st.number_input("Количество карт", min_value=1, max_value=1000, value=10, help="Укажите количество маршрутных карт для генерации")
    
    # Template file
    template_path_multi = st.text_input("Путь к шаблону PowerPoint (множественная генерация)", value="ШАБЛОН.pptx", help="Укажите путь к файлу шаблона PowerPoint")
    
    # Output directory
    output_dir_multi = st.text_input("Директория для сохранения файлов (множественная генерация)", value="Маршрутные_карты", help="Укажите директорию для сохранения сгенерированных файлов")
    
    # Generate button with custom styling
    if st.button("Создать маршрутные карты", key="multiple_generate"):
        if not os.path.exists(template_path_multi):
            st.markdown("<div class='error-box'>❌ Файл шаблона '" + template_path_multi + "' не найден!</div>", unsafe_allow_html=True)
        else:
            try:
                # Create output directory if it doesn't exist
                os.makedirs(output_dir_multi, exist_ok=True)
                
                # Generate multiple forms
                success_count, errors = generate_multiple_forms(template_path_multi, start_number, count)
                
                st.markdown("<div class='success-box'>✅ Создано " + str(success_count) + " из " + str(count) + " маршрутных карт в папке " + output_dir_multi + "</div>", unsafe_allow_html=True)
                
                if errors:
                    st.markdown("<div class='warning-box'>⚠️ Некоторые ошибки occurred:</div>", unsafe_allow_html=True)
                    for error in errors:
                        st.write(error)
                        
            except Exception as e:
                st.markdown("<div class='error-box'>❌ Ошибка при создании маршрутных карт: " + str(e) + "</div>", unsafe_allow_html=True)

# Information section in sidebar
with st.sidebar:
    st.header("ℹ️ Информация")
    st.write("Это приложение позволяет генерировать маршрутные карты в формате PowerPoint с QR-кодами.")
    st.write("Для работы приложения необходим шаблон PowerPoint с именем 'ШАБЛОН.pptx'.")
    
    st.subheader("📋 Инструкция")
    st.write("1. Убедитесь, что файл 'ШАБЛОН.pptx' находится в папке проекта")
    st.write("2. Выберите режим генерации (одна карта или несколько)")
    st.write("3. Укажите номер(а) маршрутных карт")
    st.write("4. Нажмите кнопку 'Создать'")
    
    st.subheader("📊 Статистика")
    try:
        import openpyxl
        wb = openpyxl.load_workbook('маршрутные_карты.xlsx')
        ws = wb.active
        row_count = ws.max_row - 1 if ws.max_row > 1 else 0  # Subtract 1 for header row
        st.write(f"Всего создано карт: {row_count}")
        wb.close()
    except:
        st.write("Всего создано карт: 0")