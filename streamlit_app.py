import streamlit as st
import os
import sys
from pptx import Presentation
import qrcode
from io import BytesIO
import datetime
from PIL import Image

# Add the current directory to the path so we can import our modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Импортируем модуль для работы с Google Sheets
from google_sheets_utils import save_to_google_sheets, check_duplicate_form_number, get_next_form_number, get_sheet_statistics

# Core functions moved from generate_form_pptx_core.py

def generate_form_with_qr(template_path, output_path, form_number):
    try:
        # Проверяем, существует ли уже бланк с таким номером
        if check_duplicate_form_number(form_number):
            raise ValueError(f"Бланк с номером {form_number} уже существует в базе данных")
        
        # Открываем шаблон презентации
        prs = Presentation(template_path)
        
        # Создаем QR-код с номером формы
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.ERROR_CORRECT_L,
            box_size=10,
            border=1,
        )
        qr.add_data(form_number)
        qr.make(fit=True)
        qr_image = qr.make_image(fill_color="black", back_color="white")
        
        # Сохраняем QR-код во временный буфер
        image_stream = BytesIO()
        qr_image.save(image_stream, 'PNG')
        image_stream.seek(0)
        
        # Добавляем QR-код на первый слайд
        slide = prs.slides[0]
        
        # Ищем текст "МАРШРУТНАЯ КАРТА"
        target_text = "МАРШРУТНАЯ КАРТА"
        text_shape = None
        
        for shape in slide.shapes:
            if hasattr(shape, "text") and target_text in shape.text:
                text_shape = shape
                break
        
        # Задаем размеры QR-кода
        qr_width = 500000  # ~0.5 см
        qr_height = 500000
        
        if text_shape:
            # Располагаем QR-код слева от текста на том же уровне
            left = max(300000, text_shape.left - qr_width - 200000)  # отступ от текста ~0.2 см, но не меньше 300000
            
            # Выравниваем по вертикали с текстом
            top = text_shape.top + (text_shape.height - qr_height) / 2
            
            # Добавляем QR-код на слайд
            slide.shapes.add_picture(
                image_stream,
                left,
                top,
                width=qr_width,
                height=qr_height
            )
            
            # Добавляем номер в правый верхний угол слайда
            slide_width = prs.slide_width
            number_shape = slide.shapes.add_textbox(
                slide_width - 1500000,  # отступ от правого края
                200000,  # отступ от верхнего края
                1200000,  # ширина текстового поля
                300000  # высота текстового поля
            )
            text_frame = number_shape.text_frame
            p = text_frame.paragraphs[0]
            p.text = f"№ {form_number}"
            p.alignment = 2  # выравнивание по правому краю
            
        else:
            # Если текст не найден, используем позицию по умолчанию
            left = 300000  # отступ от левого края
            top = 300000  # отступ от верхнего края
            
            # Добавляем QR-код на слайд
            slide.shapes.add_picture(
                image_stream,
                left,
                top,
                width=qr_width,
                height=qr_height
            )
            
            # Добавляем номер в правый верхний угол
            slide_width = prs.slide_width
            number_shape = slide.shapes.add_textbox(
                slide_width - 1500000,  # отступ от правого края
                200000,  # отступ от верхнего края
                1200000,  # ширина текстового поля
                300000  # высота текстового поля
            )
            text_frame = number_shape.text_frame
            p = text_frame.paragraphs[0]
            p.text = f"№ {form_number}"
            p.alignment = 2  # выравнивание по правому краю
        
        # Сохраняем результат
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        
        # Сохраняем информацию в Google Sheets
        save_to_google_sheets(form_number, output_path)
    except Exception as e:
        print(f"Ошибка при генерации формы: {e}")
        raise

def generate_multiple_forms(template_path, start_number, count):
    """Генерация нескольких маршрутных карт"""
    # Создаем папку для сохранения файлов, если она не существует
    output_dir = "Маршрутные_карты"
    os.makedirs(output_dir, exist_ok=True)
    
    # Проверяем все номера на дубликаты перед генерацией
    start_num = int(start_number)
    duplicates = []
    
    for i in range(count):
        current_num = start_num + i
        form_number = f"{current_num:06d}"
        if check_duplicate_form_number(form_number):
            duplicates.append(form_number)
    
    if duplicates:
        return 0, [f"Следующие номера бланков уже существуют в базе данных: {', '.join(duplicates)}"]
    
    # Если дубликатов нет, продолжаем генерацию
    errors = []
    success_count = 0
    
    for i in range(count):
        current_num = start_num + i
        # Форматируем номер с ведущими нулями (6 цифр)
        form_number = f"{current_num:06d}"
        output_path = os.path.join(output_dir, f"маршрутная_карта_{form_number}.pptx")
        
        try:
            generate_form_with_qr(template_path, output_path, form_number)
            success_count += 1
        except Exception as e:
            errors.append(f"Ошибка при создании файла {output_path}: {str(e)}")
    
    return success_count, errors

# Get the next form number to suggest
next_form_number = get_next_form_number()

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
        background-color: #2c3e50;
        color: #ecf0f1;
    }
    .main-header {
        background: linear-gradient(135deg, #3498db 0%, #8e44ad 100%);
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
        background-color: #34495e;
        border-radius: 8px 8px 0 0;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
        color: #ecf0f1;
    }
    .stTabs [aria-selected="true"] {
        background-color: #3498db;
        color: white;
    }
    .generate-button {
        background: linear-gradient(135deg, #3498db 0%, #8e44ad 100%);
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
        background-color: #34495e;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 20px;
        color: #ecf0f1;
    }
    .success-box {
        background-color: #27ae60;
        border-left: 5px solid #2ecc71;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
        color: #ecf0f1;
    }
    .error-box {
        background-color: #c0392b;
        border-left: 5px solid #e74c3c;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
        color: #ecf0f1;
    }
    .warning-box {
        background-color: #d35400;
        border-left: 5px solid #f39c12;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
        color: #ecf0f1;
    }
    .input-label {
        font-weight: bold;
        margin-bottom: 5px;
        display: block;
        color: #ecf0f1;
        font-size: 16px;
    }
    .input-container {
        margin-bottom: 20px;
    }
    .stTextInput > div > div > input {
        background-color: #34495e !important;
        color: #ecf0f1 !important;
        border: 1px solid #3498db !important;
    }
    .stNumberInput > div > div > input {
        background-color: #34495e !important;
        color: #ecf0f1 !important;
        border: 1px solid #3498db !important;
    }
    h1, h2, h3 {
        color: #ecf0f1;
    }
    .stMarkdown {
        color: #ecf0f1;
    }
</style>
""", unsafe_allow_html=True)

# Main header
st.markdown("<div class='main-header'><h1>📋 Генератор маршрутных карт</h1><p>Создание маршрутных карт в формате PowerPoint с QR-кодами</p></div>", unsafe_allow_html=True)

# Create tabs for single and multiple form generation
tab1, tab2 = st.tabs(["Один бланк", "Несколько бланков"])

with tab1:
    st.header("Создание одной маршрутной карты")
    
    # Information card
    st.markdown("<div class='info-card'><h3>Информация</h3><p>Следующий доступный номер маршрутной карты: <strong>" + next_form_number + "</strong></p></div>", unsafe_allow_html=True)
    
    # Input for form number with suggestion
    st.markdown("<div class='input-container'><span class='input-label'>Номер маршрутной карты</span></div>", unsafe_allow_html=True)
    form_number = st.text_input("Номер маршрутной карты", value=next_form_number, key="single_form_number", label_visibility="visible")
    
    # Template file
    st.markdown("<div class='input-container'><span class='input-label'>Путь к шаблону PowerPoint</span></div>", unsafe_allow_html=True)
    template_path = st.text_input("Путь к шаблону PowerPoint", value="ШАБЛОН.pptx", key="single_template_path", label_visibility="visible")
    
    # Output directory
    st.markdown("<div class='input-container'><span class='input-label'>Папка для сохранения файлов</span></div>", unsafe_allow_html=True)
    output_dir = st.text_input("Папка для сохранения файлов", value="Маршрутные_карты", key="single_output_dir", label_visibility="visible")
    
    # Generate button with custom styling
    if st.button("Создать маршрутную карту", key="single_generate", use_container_width=True):
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
    st.markdown("<div class='input-container'><span class='input-label'>Начальный номер</span></div>", unsafe_allow_html=True)
    start_number = st.text_input("Начальный номер", value=next_form_number, key="multiple_start_number", label_visibility="visible")
    
    # Count
    st.markdown("<div class='input-container'><span class='input-label'>Количество карт</span></div>", unsafe_allow_html=True)
    count = st.number_input("Количество карт", min_value=1, max_value=1000, value=10, key="multiple_count", label_visibility="visible")
    
    # Template file
    st.markdown("<div class='input-container'><span class='input-label'>Путь к шаблону PowerPoint</span></div>", unsafe_allow_html=True)
    template_path_multi = st.text_input("Путь к шаблону PowerPoint (множественная генерация)", value="ШАБЛОН.pptx", key="multiple_template_path", label_visibility="visible")
    
    # Output directory
    st.markdown("<div class='input-container'><span class='input-label'>Папка для сохранения файлов</span></div>", unsafe_allow_html=True)
    output_dir_multi = st.text_input("Папка для сохранения файлов (множественная генерация)", value="Маршрутные_карты", key="multiple_output_dir", label_visibility="visible")
    
    # Generate button with custom styling
    if st.button("Создать маршрутные карты", key="multiple_generate", use_container_width=True):
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
    with st.spinner("Загрузка статистики..."):
        try:
            stats = get_sheet_statistics()
            if stats:
                st.success(f"Всего карт: **{stats['total']}**")
                
                st.write("По статусам:")
                # Using a more robust way to display statuses
                status_dict = stats.get('statuses', {})
                if not status_dict:
                    st.write("Нет данных о статусах.")
                else:
                    # Create a neat layout for statuses
                    # Check if the number of statuses is not too large for columns
                    if len(status_dict) <= 5: # Arbitrary limit for using columns
                        cols = st.columns(len(status_dict))
                        for i, (status, count) in enumerate(status_dict.items()):
                            with cols[i]:
                                st.metric(label=status, value=count)
                    else:
                        for status, count in status_dict.items():
                            st.metric(label=status, value=count)
            else:
                st.warning("Не удалось загрузить статистику.")
        except Exception as e:
            st.error("Ошибка при загрузке статистики:")
            st.exception(e)
