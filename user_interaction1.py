import telebot
from telebot.types import Message, ReplyKeyboardMarkup, KeyboardButton
from appeals_handler import save_user_data
from excel_handler import save_to_excel
from logger import log_message

# Инициализация бота
TOKEN = "7662734466:AAFeIkLV6zkROToipNUU_DdBJV8ge_WRgKM"
bot = telebot.TeleBot(TOKEN)

# Хранилище состояний пользователей
user_states = {}

# Хранилище состояний пользователей
user_states = {}

# Функция для создания кнопок меню
def create_menu():
    markup = ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    buttons = [
        KeyboardButton("1. Запись к специалистам"),
        KeyboardButton("2. Телефонный справочник"),
        KeyboardButton("3. Адрес СМФЦ, геолокация"),
        KeyboardButton("4. Ссылки на сайты"),
        KeyboardButton("5. Тел. ОСЗН Тульской области"),
        KeyboardButton("6. Мероприятия СМФЦ для записи"),
        KeyboardButton("7. Консультация диспетчера")
    ]
    markup.add(*buttons)  # Добавляем все кнопки
    return markup

@bot.message_handler(commands=['start'])
def start_command(message: Message):
    """Обработчик команды /start. Запрашивает у пользователя ФИО."""
    log_message(message.from_user.id, message.from_user.username, "Command", message.text)
    save_to_excel(message.from_user.id, message.from_user.username, "Command", message.text)
    
    bot.send_message(message.chat.id, "Здравствуйте, меня зовут Дина.\nПредставьтесь, пожалуйста. (А Вас?)")
    user_states[message.chat.id] = "waiting_for_name"

@bot.message_handler(func=lambda message: user_states.get(message.chat.id) == "waiting_for_name")
def handle_user_name(message: Message):
    """Обрабатывает ФИО пользователя и сохраняет его данные."""
    full_name = message.text.strip().split()

    if len(full_name) == 3:
        last_name, first_name, middle_name = full_name
        save_user_data(message.from_user.id, message.from_user.username, last_name, first_name, middle_name, button_text=None)
        bot.send_message(message.chat.id, "Спасибо! Ваши данные записаны.")
        
        # Запрашиваем населенный пункт
        bot.send_message(message.chat.id, "Из какого города/населенного пункта Вы обращаетесь?\n(Укажите населенный пункт без указания типа населенного пункта)")
        user_states[message.chat.id] = "waiting_for_city"
    else:
        bot.send_message(message.chat.id, "Пожалуйста, введите ФИО в формате: Фамилия Имя Отчество.")

@bot.message_handler(func=lambda message: user_states.get(message.chat.id) == "waiting_for_city")
def handle_city(message: Message):
    """Обрабатывает населенный пункт пользователя и сохраняет его данные."""
    city = message.text.strip()

    # Сохраняем населенный пункт в файл
    save_user_data(message.from_user.id, message.from_user.username, city=city)

    # Отправляем меню после записи населенного пункта
    bot.send_message(message.chat.id, "Какой вопрос Вас интересует?", reply_markup=create_menu())
    user_states.pop(message.chat.id)  # Удаляем состояние

@bot.message_handler(func=lambda message: message.text in [
    "1. Запись к специалистам",
    "2. Телефонный справочник",
    "3. Адрес СМФЦ, геолокация",
    "4. Ссылки на сайты",
    "5. Тел. ОСЗН Тульской области",
    "6. Мероприятия СМФЦ для записи",
    "7. Консультация диспетчера"
])
def handle_menu_response(message: Message):
    """Обрабатывает нажатие кнопок меню, логирует и отправляет ответ."""
    button_text = message.text
    
    log_message(message.from_user.id, message.from_user.username, "Button", button_text)
    save_to_excel(message.from_user.id, message.from_user.username, "Button", message.text)
    
    save_user_data(message.from_user.id, message.from_user.username, "Фамилия", "Имя", "Отчество", button_text)  # Передаем текст кнопки

    bot.send_message(message.chat.id, "Пока данный функционал ещё не реализован.")

@bot.message_handler(func=lambda message: True)
def log_all_messages(message: Message):
    """Логирование всех текстовых сообщений пользователей."""
    log_message(message.from_user.id, message.from_user.username, "Message", message.text)
    save_to_excel(message.from_user.id, message.from_user.username, "Message", message.text)

if __name__ == "__main__":
    bot.infinity_polling()