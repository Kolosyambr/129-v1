import telebot
from telebot.types import Message, ReplyKeyboardMarkup, KeyboardButton
from appeals_handler import save_user_data, update_user_data, update_user_location, update_user_appeal, update_contact_phone

TOKEN = "7662734466:AAFeIkLV6zkROToipNUU_DdBJV8ge_WRgKM"
bot = telebot.TeleBot(TOKEN)

user_data = {}

@bot.message_handler(commands=['start'])
def handle_start(message: Message):
    user_id = message.from_user.id
    username = message.from_user.username

    # Сохраняем TG ID и Username в Excel
    save_user_data(user_id, username)

    # Приветственное сообщение
    bot.send_message(user_id, "Здравствуйте, меня зовут Дина.\nПредставьтесь, пожалуйста. (Укажите Ваше ФИО)")
    bot.register_next_step_handler(message, ask_full_name)

def ask_full_name(message: Message):
    user_id = message.from_user.id
    full_name = message.text

    # Обновляем ФИО в Excel
    update_user_data(user_id, full_name)

    bot.send_message(user_id, "Спасибо, ваши данные успешно сохранены!")

    # Запрашиваем населенный пункт
    bot.send_message(user_id, "Из какого города/населенного пункта Вы обращаетесь?\nНапишите без указания типа населенного пункта (к примеру - Тула, Новомосковск и т.д.)")
    bot.register_next_step_handler(message, ask_location)

def ask_location(message: Message):
    user_id = message.from_user.id
    location = message.text

    # Обновляем населенный пункт в Excel, если необходимо
    update_user_location(user_id, location)

    bot.send_message(user_id, "Спасибо за информацию, ваш запрос успешно обработан!")

    # Отображаем главное меню
    show_main_menu(user_id)

    # Проверка на наличие ключа перед удалением
    if user_id in user_data:
        del user_data[user_id]  # Удаляем временные данные

# Обработчик на кнопку "Запись к специалистам"
@bot.message_handler(func=lambda message: message.text in ["Запись к специалистам", "Телефонный справочник", "Адрес СМФЦ, геолокация", 
                                                            "Ссылки на сайты", "Тел. ОСЗН Тульской области", 
                                                            "Мероприятия СМФЦ для записи", "Консультация диспетчера"])
def handle_main_menu_selection(message: Message):
    user_id = message.from_user.id
    selected_option = message.text

    # Обновляем колонку "Обращение" в Excel
    update_user_appeal(user_id, selected_option)

    # Переходим к следующему шагу или показываем дополнительное меню
    if selected_option == "Запись к специалистам":
        handle_record_to_specialists(message)
    elif selected_option == "Телефонный справочник":
        handle_phone_directory(message)
    elif selected_option == "Адрес СМФЦ, геолокация":
        handle_address_and_geolocation(message)
    elif selected_option == "Ссылки на сайты":
        handle_links(message)
    elif selected_option == "Тел. ОСЗН Тульской области":
        handle_oszn_phone(message)
    elif selected_option == "Мероприятия СМФЦ для записи":
        handle_events(message)
    elif selected_option == "Консультация диспетчера":
        handle_consultation_menu(message)
    else:
        show_main_menu(user_id)

    # Если не "Запись к специалистам", обновляем контактный телефон на "Не требуется"
    if selected_option != "Запись к специалистам":
        update_contact_phone(user_id, "Не требуется")

# Обработчик для телефона справочника
@bot.message_handler(func=lambda message: message.text == "Телефонный справочник")
def handle_phone_directory(message: Message):
    user_id = message.from_user.id

    # Сообщение с телефонами
    phone_numbers = (
        "8-800-200-71-02 горячая линия Губернатора\n"
        "8-800-444-40-03 горячая линия Минздрава\n"
        "8-800-450-00-71 МФЦ (+ мобильная группа МФЦ)\n"
        "8-800-450-19-45 МФЦ (+ мобильная группа МФЦ)\n"
        "8-800-100-00-01 СФР\n"
        "8-800-350-67-71 Центр занятости населения\n"
        "8-800-200-52-26 Центр системного долговременного ухода за гражданами пожилого возраста и инвалидами по мобилизации (+соцтакси)\n"
        "8-800-737-77-37 ЕРЦ Минобороны\n"
    )
    
    bot.send_message(user_id, phone_numbers)

    # После отправки телефона возвращаем пользователя в главное меню
    show_main_menu(user_id)

# Функции для различных опций в меню

def handle_record_to_specialists(message: Message):
    user_id = message.from_user.id

    # Создаем клавиатуру для записи к специалистам
    markup = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(KeyboardButton("Юрист"))
    markup.add(KeyboardButton("Психолог"))
    markup.add(KeyboardButton("Служба Экстренного Реагирования"))
    markup.add(KeyboardButton("Соцзащита"))
    markup.add(KeyboardButton("Назад в Главное меню"))

    # Отправляем меню с кнопками
    bot.send_message(user_id, "Выберите специалиста:", reply_markup=markup)

    # Сохраняем выбранного специалиста в словарь для дальнейшего использования
    user_data[user_id] = {"specialist": message.text}

def handle_address_and_geolocation(message: Message):
    user_id = message.from_user.id
    # Пример обработки запроса на геолокацию или адрес СМФЦ
    bot.send_message(user_id, "Здесь можно добавить информацию по адресу СМФЦ и геолокации.")

def handle_links(message: Message):
    user_id = message.from_user.id

    # линки
    links = (
        "Правительство ТО: https://tularegion.ru/\n"
        "Министерство труда и социальной защиты ТО: https://mintrud.tularegion.ru/about/contacts/\n"
        "Управление социальной защиты ТО: https://tulauszn.gosuslugi.ru/services/\n"
        "СМФЦ «Мой семейный центр» в Телеграм-канале https://t.me/msc_tula\n"
    )
    
    bot.send_message(user_id, links)

    # После отправки телефона возвращаем пользователя в главное меню
    show_main_menu(user_id)

def handle_oszn_phone(message: Message):
    user_id = message.from_user.id
    # Пример обработки запроса на телефон ОСЗН Тульской области
    bot.send_message(user_id, "Телефон ОСЗН Тульской области: 8-800-100-00-01.")

def handle_events(message: Message):
    user_id = message.from_user.id
    # Пример обработки запроса на мероприятия СМФЦ
    bot.send_message(user_id, "Здесь можно добавить информацию о мероприятиях для записи.")

def handle_consultation_menu(message: Message):
    user_id = message.from_user.id

    # Создаем клавиатуру для консультации диспетчера
    markup = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(KeyboardButton("Семьи с детьми"))
    markup.add(KeyboardButton("Государственная Социальная помощь"))
    markup.add(KeyboardButton("Социальный контракт"))
    markup.add(KeyboardButton("Участники СВО и члены их семей"))
    markup.add(KeyboardButton("Прочее"))
    markup.add(KeyboardButton("Назад в Главное меню"))

    # Отправляем меню с кнопками
    bot.send_message(user_id, "Выберите, по какому вопросу Вам нужна консультация:", reply_markup=markup)

# Функции для обработки выбора из меню консультации диспетчера
@bot.message_handler(func=lambda message: message.text in ["Семьи с детьми", "Государственная Социальная помощь", 
                                                           "Социальный контракт", "Участники СВО и члены их семей", 
                                                           "Прочее"])
def handle_consultation_option(message: Message):
    user_id = message.from_user.id
    selected_option = message.text

    # Здесь можно добавить логику для каждой из выбранных категорий
    if selected_option == "Семьи с детьми":
        bot.send_message(user_id, "Консультация по теме 'Семьи с детьми': Подробная информация будет предоставлена.")
    elif selected_option == "Государственная Социальная помощь":
        bot.send_message(user_id, "Консультация по теме 'Государственная Социальная помощь': Подробная информация будет предоставлена.")
    elif selected_option == "Социальный контракт":
        bot.send_message(user_id, "Консультация по теме 'Социальный контракт': Подробная информация будет предоставлена.")
    elif selected_option == "Участники СВО и члены их семей":
        bot.send_message(user_id, "Консультация по теме 'Участники СВО и члены их семей': Подробная информация будет предоставлена.")
    elif selected_option == "Прочее":
        bot.send_message(user_id, "Консультация по теме 'Прочее': Подробная информация будет предоставлена.")

    # Возвращаем пользователя в главное меню
    show_main_menu(user_id)

# Обработчик для ввода номера телефона
@bot.message_handler(func=lambda message: message.text in ["Юрист", "Психолог", "Служба Экстренного Реагирования", "Соцзащита"])
def handle_specialist_selection(message: Message):
    user_id = message.from_user.id
    specialist = message.text  # Получаем выбранного специалиста

    # Просим пользователя указать номер телефона
    bot.send_message(user_id, f"Укажите пожалуйста, Ваш актуальный номер телефона и в ближайшее время с Вами свяжется диспетчер.")

    # Сохраняем выбранного специалиста, чтобы потом обработать номер телефона
    user_data[user_id] = {"specialist": specialist}

    # Переход к шагу обработки номера телефона
    bot.register_next_step_handler(message, handle_phone_number)

# Обработчик для ввода номера телефона
def handle_phone_number(message: Message):
    user_id = message.from_user.id
    phone_number = message.text

    # Сохраняем номер телефона в соответствующей строке Excel
    update_contact_phone(user_id, phone_number)

    # Подтверждение получения номера телефона
    bot.send_message(user_id, f"Ваш номер телефона {phone_number} записан. В ближайшее время с Вами свяжется диспетчер.")

    # Возвращаем пользователя в главное меню
    show_main_menu(user_id)

# Обработчик на кнопку "Назад в Главное меню"
@bot.message_handler(func=lambda message: message.text == "Назад в Главное меню")
def handle_back_to_main_menu(message: Message):
    user_id = message.from_user.id

    # Отправляем главное меню
    show_main_menu(user_id)

# Функция для отображения главного меню
def show_main_menu(user_id):
    markup = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(KeyboardButton("Запись к специалистам"))
    markup.add(KeyboardButton("Телефонный справочник"))
    markup.add(KeyboardButton("Адрес СМФЦ, геолокация"))
    markup.add(KeyboardButton("Ссылки на сайты"))
    markup.add(KeyboardButton("Тел. ОСЗН Тульской области"))
    markup.add(KeyboardButton("Мероприятия СМФЦ для записи"))
    markup.add(KeyboardButton("Консультация диспетчера"))

    bot.send_message(user_id, "Какой вопрос Вас интересует?", reply_markup=markup)

if __name__ == "__main__":
    bot.infinity_polling()