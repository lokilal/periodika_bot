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
        self.max_row = self.sheet.max_row

    def get_data(self):
        data = []
        for row in self.sheet['B2':f'C{self.max_row}']:
            data_from_fow = []
            for cell in row:
                data_from_fow.append(cell.value)
            data.append(data_from_fow)
        return data

    def get_id(self):
        data = []
        for row in self.sheet['A2':f'A{self.max_row}']:
            for cell in row:
                data.append(cell.value)
        return data

    def write_new_user(self, telegram_id, username):
        self.sheet[f'A{self.max_row} + 1'] = telegram_id
        self.sheet[f'B{self.max_row}'] = username
        self.workbook.save(PATH_TO_EXCEL)

    def reload_table(self):
        self.__init__()


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
        bot.send_message(
            message.chat.id, 'Таблица перезагружена'
        )


@bot.message_handler(content_types='text')
def message_reply(message):
    if message.text == 'Получить промокод':
        print(table.get_data())
    elif message.text == 'Получить баланс':
        pass
    elif message.text == 'Потратить':
        pass
    else:
        bot.send_message(
            message.chat.id, 'Я вас не понял'
        )


table = ExcelTable()
bot.infinity_polling()
