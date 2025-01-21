import os
import pandas as pd
import telebot
from telebot import types
import re
import time

# Константы
BOT_TOKEN = '7247840603:AAEHxB8Mv2siQKIKIdxi3Rs5eR9lyziUtwA'
ADMIN_ID = 374805661
BASE_PATH = r'C:\kislorod_base'
BASE_FILE = os.path.join(BASE_PATH, 'car_registry.xlsx')

# Создание папки, если не существует
os.makedirs(BASE_PATH, exist_ok=True)

# Инициализация бота
bot = telebot.TeleBot(BOT_TOKEN)

# Словарь для хранения состояний пользователей
user_states = {}

# Проверка корректности номера телефона
def validate_phone(phone):
    return re.match(r'^8\d{10}$', phone)

# Проверка корректности номера автомобиля
def validate_car_number(car_number):
    return len(car_number) == 6  # Проверка только на длину

# Загрузка базы данных
def load_database():
    try:
        df = pd.read_excel(BASE_FILE)
        # Проверка наличия необходимых столбцов
        required_columns = ['Телефон', 'Имя', 'Фамилия', 'Корпус', 'Квартира', 'Марка1', 'Номер1', 'Регион1', 'Марка2', 'Номер2', 'Регион2', 'Марка3', 'Номер3', 'Регион3']
        for column in required_columns:
            if column not in df.columns:
                df[column] = None  # Добавляем отсутствующий столбец с пустыми значениями
        return df
    except FileNotFoundError:
        return pd.DataFrame(columns=['Телефон', 'Имя', 'Фамилия', 'Корпус', 'Квартира', 'Марка1', 'Номер1', 'Регион1', 'Марка2', 'Номер2', 'Регион2', 'Марка3', 'Номер3', 'Регион3'])

# Сохранение базы данных
def save_database(df):
    df.to_excel(BASE_FILE, index=False)

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start_message(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    reg_button = types.KeyboardButton('Регистрация')
    # Добавляем кнопку поиска и просмотра всех зарегистрировавшихся для администратора
    if message.from_user.id == ADMIN_ID:
        search_button = types.KeyboardButton('Поиск')
        view_all_button = types.KeyboardButton('Показать всех зарегистрировавшихся')
        send_excel_button = types.KeyboardButton('Отправить Excel')  # Новая кнопка
        markup.add(search_button, view_all_button, send_excel_button)  # Добавляем кнопку в разметку
    markup.add(reg_button)

    bot.send_message(message.chat.id,
                     "Здравствуйте. Я автобот ж/к \"Кислород\". Пожалуйста, зарегистрируйтесь в базе нашего дома.",
                     reply_markup=markup)

# Обработчик показа всех зарегистрировавшихся
@bot.message_handler(func=lambda message: message.text == 'Показать всех зарегистрировавшихся' and message.from_user.id == ADMIN_ID)
def show_all_registered(message):
    df = load_database()
    if df.empty:
        bot.send_message(message.chat.id, "Нет зарегистрированных пользователей.")
    else:
        # Подсчет количества машин по корпусам
        car_counts = df['Корпус'].value_counts()
        response = "Количество машин по корпусам:\n"
        for corpus, count in car_counts.items():
            response += f"Корпус {corpus}: {count} машин(ы)\n"
        bot.send_message(message.chat.id, response)

        # Разбиваем сообщение на части, если оно слишком длинное
        max_length = 4096
        while len(response) > max_length:
            # Находим последний пробел в пределах максимальной длины
            split_index = response.rfind(' ', 0, max_length)
            if split_index == -1:  # Если пробел не найден, просто обрезаем
                split_index = max_length

            # Отправляем часть сообщения
            bot.send_message(message.chat.id, response[:split_index])
            # Удаляем отправленную часть из полного сообщения
            response = response[split_index:].strip()

        # Отправляем оставшуюся часть сообщения
        if response:
            bot.send_message(message.chat.id, response)

# Обработчик отправки Excel файла
@bot.message_handler(func=lambda message: message.text == 'Отправить Excel' and message.from_user.id == ADMIN_ID)
def send_excel_file(message):
    df = load_database()
    if df.empty:
        bot.send_message(message.chat.id, "Нет зарегистрированных пользователей для отправки.")
    else:
        # Отправляем файл
        with pd.ExcelWriter(BASE_FILE, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Регистрация')
        with open(BASE_FILE, 'rb') as file:
            bot.send_document(message.chat.id, file, caption="Вот файл с данными всех зарегистрированных пользователей.")

# Обработчик поиска данных по номеру автомобиля
@bot.message_handler(func=lambda message: message.text == 'Поиск' and message.from_user.id == ADMIN_ID)
def search_car(message):
    bot.send_message(message.chat.id, "Введите номер автомобиля для поиска:")
    user_states[message.chat.id] = {'state': 'searching_car'}

@bot.message_handler(func=lambda message: user_states.get(message.chat.id, {}).get('state') == 'searching_car')
def handle_search(message):
    df = load_database()
    car_number = message.text.strip().upper()
    result = df[df['Номер1'] == car_number]

    if not result.empty:
        info = result.iloc[0]
        response = f"""
Номер телефона: {info.get('Телефон', 'Не указано')}
Корпус: {info.get('Корпус', 'Не указано')}
Квартира: {info.get('Квартира', 'Не указана')}
Имя: {info.get('Имя', 'Не указано')}
Фамилия: {info.get('Фамилия', 'Не указана')}
Марка автомобиля: {info.get('Марка1', 'Не указана')}
Номер автомобиля: {info.get('Номер1', 'Не указан')}
Регион: {info.get('Регион1', 'Не указан')} 
"""
        bot.send_message(message.chat.id, response)
    else:
        bot.send_message(message.chat.id, "Автомобиль не найден.")

    user_states[message.chat.id] = {}  # Сброс состояния

# Основной обработчик сообщений
@bot.message_handler(func=lambda message: True)
def handle_message(message):
    chat_id = message.chat.id

    # Начало регистрации
    if message.text == 'Регистрация':
        bot.send_message(chat_id, "Введите имя:")
        user_states[chat_id] = {'state': 'entering_name', 'car_count': 1}
        return

    # Получаем текущее состояние
    state = user_states.get(chat_id, {})
    current_state = state.get('state')

    if current_state == 'entering_name':
        state['name'] = message.text.strip()
        bot.send_message(chat_id, "Введите фамилию:")
        state['state'] = 'entering_surname'

    elif current_state == 'entering_surname':
        state['surname'] = message.text.strip()
        bot.send_message(chat_id, "Введите номер корпуса (1 или 2):")
        state['state'] = 'entering_corpus'

    elif current_state == 'entering_corpus':
        if message.text in ['1', '2']:
            state['corpus'] = message.text
            bot.send_message(chat_id, "Введите номер квартиры:")  # Новый ввод номера квартиры
            state['state'] = 'entering_apartment'  # Обновляем состояние
        else:
            bot.send_message(chat_id, "Пожалуйста, введите 1 или 2:")

    # Новый блок для ввода номера квартиры
    elif current_state == 'entering_apartment':
        state['apartment'] = message.text.strip()  # Сохраняем номер квартиры
        bot.send_message(chat_id, "Введите номер телефона (8XXXXXXXXXX):")
        state['state'] = 'entering_phone'

    elif current_state == 'entering_phone':
        if validate_phone(message.text):
            state['phone'] = message.text
            bot.send_message(chat_id, "Введите марку автомобиля:")
            state['state'] = 'entering_car_brand'
        else:
            bot.send_message(chat_id, "Некорректный номер телефона. Введите в формате 8XXXXXXXXXX:")

    elif current_state == 'entering_car_brand':
        state[f'car{state["car_count"]}_brand'] = message.text.strip()
        bot.send_message(chat_id, "Введите номер автомобиля (6 символов):")
        state['state'] = 'entering_car_number'

    elif current_state == 'entering_car_number':
        car_number_input = message.text.strip().upper()  # Переводим номер в верхний регистр
        if validate_car_number(car_number_input):  # Проверка только на длину
            state[f'car{state["car_count"]}_number'] = car_number_input  # Сохраняем номер с измененным регистром
            bot.send_message(chat_id, "Введите номер региона (2 или 3 цифры):")
            state['state'] = 'entering_region'
        else:
            bot.send_message(chat_id, "Некорректный номер автомобиля. Он должен содержать 6 символов (буквы и цифры).")

    elif current_state == 'entering_region':
        if message.text.isdigit() and (2 <= len(message.text) <= 3):
            state[f'car{state["car_count"]}_region'] = message.text.strip()
            # Кнопка "Проверить и сохранить"
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            check_button = types.KeyboardButton('Проверить и сохранить')
            markup.add(check_button)

            bot.send_message(chat_id, "Данные введены. Нажмите 'Проверить и сохранить'.", reply_markup=markup)
            state['state'] = 'pre_confirmation'
        else:
            bot.send_message(chat_id, "Некорректный номер региона. Введите 2 или 3 цифры:")

    elif current_state == 'pre_confirmation' and message.text == 'Проверить и сохранить':
        # Вывод подтверждения
        confirmation_text = f"""
Проверьте введенные данные:
Имя: {state['name']}
Фамилия: {state['surname']}
Корпус: {state['corpus']}
Квартира: {state['apartment']}
Телефон: {state['phone']}
Марка авто: {state[f'car{state["car_count"]}_brand']}
Номер авто: {state[f'car{state["car_count"]}_number']}
Регион: {state[f'car{state["car_count"]}_region']}
"""
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        confirm_button = types.KeyboardButton('Всё верно')
        edit_button = types.KeyboardButton('Исправить')
        markup.add(confirm_button, edit_button)

        bot.send_message(chat_id, confirmation_text, reply_markup=markup)
        state['state'] = 'confirmation'

    elif current_state == 'confirmation':
        if message.text == 'Всё верно':
            # Сохранение данных
            save_user_data(chat_id)

            # Кнопки после сохранения
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            if state['car_count'] == 1:
                second_car_button = types.KeyboardButton('Зарегистрировать второй автомобиль')
                markup.add(second_car_button)
            elif state['car_count'] == 2:
                third_car_button = types.KeyboardButton('Зарегистрировать третий автомобиль')
                markup.add(third_car_button)

            finish_button = types.KeyboardButton('Завершить')
            markup.add(finish_button)

            bot.send_message(chat_id, "Спасибо, Ваш автомобиль зарегистрирован!", reply_markup=markup)
            state['state'] = 'post_registration'

        elif message.text == 'Исправить':
            # Возврат к началу регистрации
            start_message(message)
            user_states[chat_id] = None

    elif current_state == 'post_registration':
        if message.text == 'Зарегистрировать второй автомобиль':
            state['car_count'] = 2
            bot.send_message(chat_id, "Введите марку второго автомобиля:")
            state['state'] = 'entering_car_brand'

        elif message.text == 'Зарегистрировать третий автомобиль':
            state['car_count'] = 3
            bot.send_message(chat_id, "Введите марку третьего автомобиля:")
            state['state'] = 'entering_car_brand'

        elif message.text == 'Завершить':
            bot.send_message(chat_id, "Спасибо, Ваш автомобиль зарегистрирован!")
            bot.send_message(chat_id, "Благодарю за регистрацию Вашего автомобиля!\n"
                                       "Все вопросы и предложения по работе бота @LexxDramma")
            user_states[chat_id] = None
            # Удаляем состояние, чтобы не возвращаться к регистрации

    # Обновляем состояние пользователя
    user_states[chat_id] = state

# Сохранение данных пользователя
def save_user_data(chat_id):
    state = user_states[chat_id]
    df = load_database()

    # Проверка на существование записи с таким телефоном
    existing = df[df['Телефон'] == state['phone']]

    new_entry = {
        'Телефон': state['phone'],
        'Имя': state['name'],
        'Фамилия': state['surname'],
        'Корпус': state['corpus'],
        'Квартира': state['apartment'],  # Добавлено сохранение номера квартиры
        'Марка1': state['car1_brand'],
        'Номер1': state['car1_number'],
        'Регион1': state['car1_region']
    }

    # Добавление второго авто, если есть
    if state.get('car2_brand'):
        new_entry.update({
            'Марка2': state['car2_brand'],
            'Номер2': state['car2_number'],
            'Регион2': state['car2_region']
        })

    # Добавление третьего авто, если есть
    if state.get('car3_brand'):
        new_entry.update({
            'Марка3': state['car3_brand'],
            'Номер3': state['car3_number'],
            'Регион3': state['car3_region']
        })

    if existing.empty:
        # Используем pd.concat вместо append
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
    else:
        df.loc[df['Телефон'] == state['phone']] = new_entry

    save_database(df)

# Запуск бота
while True:
    try:
        bot.polling(none_stop=True)
    except Exception as e:
        print(f"Произошла ошибка: {e}. Перезапуск бота через 5 секунд...")
        time.sleep(5)  # Задержка перед перезапуском