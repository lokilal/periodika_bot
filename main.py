import os

import openpyxl
import telebot
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv('TOKEN')
PATH_TO_EXCEL = os.getenv('PATH_TO_EXCEL')
WORKBOOK = openpyxl.load_workbook(PATH_TO_EXCEL)
SHEET = WORKBOOK['main']

bot = telebot.TeleBot(TOKEN)


def get_data_from_excel():
    data = []
    for row in SHEET['B2':f'C{SHEET.max_row}']:
        data_from_fow = []
        for cell in row:
            data_from_fow.append(cell.value)
        data.append(data_from_fow)
    return data


def get_id_from_excel():
    data = []
    for row in SHEET['A2':f'A{SHEET.max_row}']:
        for cell in row:
            data.append(cell.value)
    return data


@bot.message_handler(commands=['start'])
def start_message(message):
    if message.chat.id not in get_id_from_excel():
        SHEET[f'A{SHEET.max_row + 1}'] = message.chat.id
        SHEET[f'B{SHEET.max_row }'] = message.chat.username
        WORKBOOK.save(PATH_TO_EXCEL)
    bot.send_message(message.chat.id, 'Hello World')


bot.infinity_polling()
