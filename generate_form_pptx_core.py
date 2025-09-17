from pptx import Presentation
import qrcode
import os
from io import BytesIO
import datetime
from openpyxl import Workbook, load_workbook

def create_database():
    """Создание Excel файла, если он не существует"""
    excel_file = 'маршрутные_карты.xlsx'
    if not os.path.exists(excel_file):
        wb = Workbook()
        ws = wb.active
        ws.title = "Маршрутные карты"
        
        # Создаем заголовки столбцов
        headers = ['id', 'Номер_бланка', 'Учетный_номер', 'Номер_кластера', 'Статус', 'Дата_создания', 'Путь_к_файлу']
        for col_num, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_num, value=header)
        
        # Сохраняем файл
        wb.save(excel_file)

def save_to_database(form_number, file_path):
    """Сохранение информации о созданной маршрутной карте в Excel файл"""
    try:
        excel_file = 'маршрутные_карты.xlsx'
        
        # Загружаем существующий файл или создаем новый
        if os.path.exists(excel_file):
            wb = load_workbook(excel_file)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Маршрутные карты"
            
            # Создаем заголовки столбцов
            headers = ['id', 'Номер_бланка', 'Учетный_номер', 'Номер_кластера', 'Статус', 'Дата_создания', 'Путь_к_файлу']
            for col_num, header in enumerate(headers, 1):
                ws.cell(row=1, column=col_num, value=header)
        
        # Определяем следующий ID
        max_id = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is not None:
                max_id = max(max_id, row[0])
        
        next_id = max_id + 1
        date_created = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Добавляем новую строку с данными
        new_row = [next_id, form_number, None, None, None, date_created, file_path]
        ws.append(new_row)
        
        # Сохраняем файл
        wb.save(excel_file)
    except Exception as e:
        print(f"Ошибка при сохранении в Excel файл: {e}")
        raise

def check_duplicate_form_number(form_number):
    """Проверка существования бланка с таким номером в Excel файле"""
    excel_file = 'маршрутные_карты.xlsx'
    
    # Если файл не существует, дубликатов нет
    if not os.path.exists(excel_file):
        return False
    
    try:
        wb = load_workbook(excel_file)
        ws = wb.active
        
        # Проверяем каждую строку, начиная со второй (первая - заголовки)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[1] == form_number:  # row[1] соответствует столбцу 'Номер_бланка'
                return True
        
        return False
    except Exception as e:
        print(f"Ошибка при проверке дубликатов в Excel файле: {e}")
        return False

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
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=1,
        )
        qr.add_data(form_number)
        qr.make(fit=True)
        qr_image = qr.make_image(fill_color="black", back_color="white")
        
        # Сохраняем QR-код во временный буфер
        image_stream = BytesIO()
        qr_image.save(image_stream, format='PNG')
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
        
        # Сохраняем информацию в Excel файл
        save_to_database(form_number, output_path)
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

def update_empty_to_null():
    """Обновление пустых значений на NULL в существующем Excel файле"""
    excel_file = 'маршрутные_карты.xlsx'
    
    # Если файл не существует, ничего не делаем
    if not os.path.exists(excel_file):
        return
    
    try:
        wb = load_workbook(excel_file)
        ws = wb.active
        
        # Проходим по всем строкам, начиная со второй (первая - заголовки)
        for row_num in range(2, ws.max_row + 1):
            # Проверяем столбцы: Учетный_номер (C), Номер_кластера (D), Статус (E)
            # В Excel: C=3, D=4, E=5
            for col_num in [3, 4, 5]:
                cell_value = ws.cell(row=row_num, column=col_num).value
                if cell_value == "":
                    ws.cell(row=row_num, column=col_num).value = None
        
        # Сохраняем изменения
        wb.save(excel_file)
    except Exception as e:
        print(f"Ошибка при обновлении Excel файла: {e}")
        raise

def get_next_form_number():
    """Получение следующего доступного номера маршрутной карты"""
    excel_file = 'маршрутные_карты.xlsx'
    
    # Если файл не существует, возвращаем начальный номер
    if not os.path.exists(excel_file):
        return "000001"
    
    try:
        wb = load_workbook(excel_file)
        ws = wb.active
        
        # Если нет данных, возвращаем начальный номер
        if ws.max_row <= 1:
            return "000001"
        
        # Получаем все номера бланков
        form_numbers = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[1] is not None:  # row[1] соответствует столбцу 'Номер_бланка'
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
                # Если не удалось преобразовать в число, пропускаем
                continue
        
        next_number = max_number + 1
        return f"{next_number:06d}"
    except Exception as e:
        print(f"Ошибка при получении следующего номера: {e}")
        return "000001"

def main():
    create_database()  # Сначала создаем таблицу
    try:
        update_empty_to_null()  # Затем пытаемся обновить значения
    except Exception as e:
        print(f"Предупреждение при обновлении Excel файла: {e}")
        # Продолжаем выполнение даже при ошибке
    print("Основные функции доступны. Используйте generate_form_with_qr() или generate_multiple_forms() для создания маршрутных карт.")

if __name__ == "__main__":
    main()