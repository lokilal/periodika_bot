import os

import openpyxl
import telebot
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv('TOKEN')
PATH_TO_EXCEL = os.getenv('PATH_TO_EXCEL')
ADMINS = os.getenv('ADMINS').split(',')
BUTTONS = ['Получить промокод', 'Получить баланс', 'Потратить']
bot = telebot.TeleBot(TOKEN)


class ExcelTable:
    def __init__(self):
        self.workbook = openpyxl.load_workbook(PATH_TO_EXCEL)
        self.sheet = self.workbook['main']

    def get_user_data(self, telegram_id: int) -> list:
        """
        Возвращает данные пользователя.
        :param telegram_id: Уникальный id пользователя
        :return: Словарь с данными пользователя
        """
        row_number = self.get_id().index(telegram_id) + 2
        data = []
        for column in self.sheet[f'B{row_number}':f'E{row_number}']:
            for cell in column:
                data.append(cell.value)
        return data

    def get_id(self) -> list:
        """
        Возвращает список id всех пользователей.
        :return: Список id всех пользователей
        """
        data = []
        for row in self.sheet['A2':'A1000']:
            for cell in row:
                if cell.value is not None:
                    data.append(cell.value)
        return data

    def write_new_user(self, telegram_id: int, username: str):
        """
        Записывает данные нового пользователя в таблицу.
        :param telegram_id: ID Пользователя
        :param username: Никнейм пользователя
        """
        line_number = len(self.get_id()) + 2
        self.sheet[f'A{line_number}'] = telegram_id
        self.sheet[f'B{line_number}'] = username
        self.workbook.save(PATH_TO_EXCEL)

    def delete_usage_and_result_promocode(self, telegram_id: int):
        """
        Удаляет данные из колонок usage, result_promocode.
        :param telegram_id: Уникальный ID пользователя
        """
        row_number = self.get_id().index(telegram_id) + 2
        self.sheet[f'D{row_number}'] = 0
        self.sheet[f'E{row_number}'] = 0
        self.workbook.save(PATH_TO_EXCEL)

    def reload_table(self):
        """
        Перезагружает таблицу, для обновления данных.
        Отправляет уведомление пользователю о том, что он может использовать
        промокод.
        """
        self.__init__()
        for user_id in self.get_id():
            user_data = self.get_user_data(user_id)
            if user_data[3] is not None and user_data[3] != 0:
                bot.send_message(
                    user_id,
                    'Вы можете потратить свой баланс используя промокод'
                )


@bot.message_handler(commands=['start'])
def start_message(message):
    if message.chat.id not in table.get_id():
        table.write_new_user(
            message.chat.id, message.chat.username
        )
    markup = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)
    for button in BUTTONS:
        markup.add(telebot.types.KeyboardButton(button))
    bot.send_message(
        message.chat.id, 'Доброго времени суток',
        reply_markup=markup
    )


@bot.message_handler(commands=['reload'])
def reload_table_from_message(message):
    if str(message.chat.id) in ADMINS:
        table.reload_table()
        bot.send_message(
            message.chat.id, 'Таблица перезагружена'
        )


@bot.message_handler(content_types='text')
def message_reply(message):
    user_id = message.chat.id
    have_user_data = False
    data = [None for i in range(4)]
    if user_id in table.get_id():
        data = table.get_user_data(user_id)
        have_user_data = True
    if message.text == 'Получить промокод' and have_user_data:
        bot.send_message(
            message.chat.id,
            f'Ваш промокод:\n {data[1]}'
        )
    elif message.text == 'Получить баланс' and have_user_data:
        usage = ('Обновляем данные покупок раз в неделю.Если есть вопросы, '
                 'можете обратиться в службу поддержки')
        if data[2] != 0 and data[2] is not None:
            usage = data[2]
        bot.send_message(
            user_id,
            f'Ваш баланс: \n{usage}'
        )
    elif message.text == 'Потратить' and have_user_data:
        result_promocode = ('Обновляем данные покупок раз в месяц.'
                            'Если есть вопросы, '
                            'можете обратиться в службу поддержки')
        if data[3] != 0 and data[3] is not None:
            result_promocode = data[3]
        bot.send_message(
            user_id,
            f'Промокод: \n{result_promocode}'
        )
        table.delete_usage_and_result_promocode(user_id)
    else:
        bot.send_message(
            message.chat.id, 'Я вас не понял'
        )


table = ExcelTable()
bot.infinity_polling()
