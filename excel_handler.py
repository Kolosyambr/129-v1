import openpyxl
from datetime import datetime

FILE_NAME = "User_Messages.xlsx"

def save_to_excel(user_id, username, message_type, text):
    """Сохраняет сообщения пользователей в Excel-файл с добавлением даты и времени."""
    try:
        wb = openpyxl.load_workbook(FILE_NAME)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Дата и время", "TG ID", "Username", "Тип сообщения", "Текст сообщения"])  # Заголовки
    else:
        sheet = wb.active

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Получаем текущее время в формате ГГГГ-ММ-ДД ЧЧ:ММ:СС
    sheet.append([current_time, user_id, username, message_type, text])

    wb.save(FILE_NAME)