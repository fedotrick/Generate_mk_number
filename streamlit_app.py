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
    generate_multiple_forms
)

# Initialize the database
create_database()
try:
    update_empty_to_null()
except Exception as e:
    st.warning(f"Предупреждение при обновлении базы данных: {e}")

# Streamlit app
st.title("Генератор маршрутных карт")

# Create tabs for single and multiple form generation
tab1, tab2 = st.tabs(["Один бланк", "Несколько бланков"])

with tab1:
    st.header("Создание одной маршрутной карты")
    
    # Input for form number
    form_number = st.text_input("Номер маршрутной карты", value="000001")
    
    # Template file
    template_path = st.text_input("Путь к шаблону PowerPoint", value="ШАБЛОН.pptx")
    
    # Output directory
    output_dir = st.text_input("Директория для сохранения файлов", value="Маршрутные_карты")
    
    # Generate button
    if st.button("Создать маршрутную карту"):
        if not os.path.exists(template_path):
            st.error(f"Файл шаблона '{template_path}' не найден!")
        else:
            try:
                # Create output directory if it doesn't exist
                os.makedirs(output_dir, exist_ok=True)
                
                # Format the form number
                form_number_formatted = f"{int(form_number):06d}"
                output_path = os.path.join(output_dir, f"маршрутная_карта_{form_number_formatted}.pptx")
                
                # Generate the form
                generate_form_with_qr(template_path, output_path, form_number_formatted)
                st.success(f"Маршрутная карта успешно создана: {output_path}")
                
                # Provide download link
                with open(output_path, "rb") as file:
                    st.download_button(
                        label="Скачать файл",
                        data=file,
                        file_name=f"маршрутная_карта_{form_number_formatted}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"Ошибка при создании маршрутной карты: {e}")

with tab2:
    st.header("Создание нескольких маршрутных карт")
    
    # Starting number
    start_number = st.text_input("Начальный номер", value="100001")
    
    # Count
    count = st.number_input("Количество карт", min_value=1, max_value=1000, value=10)
    
    # Template file
    template_path_multi = st.text_input("Путь к шаблону PowerPoint (множественная генерация)", value="ШАБЛОН.pptx")
    
    # Output directory
    output_dir_multi = st.text_input("Директория для сохранения файлов (множественная генерация)", value="Маршрутные_карты")
    
    # Generate button
    if st.button("Создать маршрутные карты"):
        if not os.path.exists(template_path_multi):
            st.error(f"Файл шаблона '{template_path_multi}' не найден!")
        else:
            try:
                # Create output directory if it doesn't exist
                os.makedirs(output_dir_multi, exist_ok=True)
                
                # Generate multiple forms
                success_count, errors = generate_multiple_forms(template_path_multi, start_number, count)
                
                st.success(f"Создано {success_count} из {count} маршрутных карт в папке {output_dir_multi}")
                
                if errors:
                    st.warning("Некоторые ошибки occurred:")
                    for error in errors:
                        st.write(error)
                        
            except Exception as e:
                st.error(f"Ошибка при создании маршрутных карт: {e}")

# Information section
st.sidebar.header("Информация")
st.sidebar.write("Это приложение позволяет генерировать маршрутные карты в формате PowerPoint с QR-кодами.")
st.sidebar.write("Для работы приложения необходим шаблон PowerPoint с именем 'ШАБЛОН.pptx'.")