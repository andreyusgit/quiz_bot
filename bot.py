import logging
import schedule
import time
import openpyxl
from openpyxl import Workbook

from aiogram import Bot, types
from aiogram.utils import executor
from aiogram.dispatcher import Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.contrib.middlewares.logging import LoggingMiddleware

from config import TOKEN
from utils import TestStates
from messages import MESSAGES

logging.basicConfig(format=u'%(filename)+13s [ LINE:%(lineno)-4s] %(levelname)-8s [%(asctime)s] %(message)s',
                    level=logging.DEBUG)

bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())

ws = openpyxl.load_workbook('ques1.xlsx')
sh1 = ws['Sheet1']
sh2 = ws['Sheet2']

users = {}
score = {}
qs = {}


@dp.message_handler(commands=['start'])
async def process_start_command(message: types.Message):
    state = dp.current_state(user=message.from_user.id)
    await message.answer(MESSAGES['start'])
    users[int(message.from_user.id)] = len(users)
    score[int(message.from_user.id)] = 0
    qs[int(message.from_user.id)] = 1

    wb = openpyxl.load_workbook('users.xlsx')
    sheet = wb['users']
    sheet = wb.active
    cell = sheet.cell(row=users[int(message.from_user.id)] + 2, column=1)
    cell.value = message.from_user.id
    cell = sheet.cell(row=users[int(message.from_user.id)] + 2, column=3)
    cell.value = 0
    cell = sheet.cell(row=users[int(message.from_user.id)] + 2, column=5)
    cell.value = 1
    wb.save('users.xlsx')
    wb.close()
    await state.set_state(TestStates.all()[int(6)])


@dp.message_handler(state='*', commands=['info'])
async def process_help_command(message: types.Message):
    await message.answer(MESSAGES['info'])


@dp.message_handler(state='*', commands=['thanks'])
async def process_thx_command(message: types.Message):
    await message.answer(MESSAGES['thx'])


@dp.message_handler(state=TestStates.TEST_STATE_2 | TestStates.TEST_STATE_0, commands=['quiz'])
async def process_quiz_command(message: types.Message):
    if len(qs) == 0:
        await message.answer('Викторина начинается!\n\n'
                             'Вводи ответы развернуто и называй фильмы/сериалы '
                             'их полным названием с указанием части\n\nУспехов !')
    if qs[int(message.from_user.id)] > 20:
        await message.answer('Введите ответ на 1 вопрос')
    else:
        await message.answer('Введите ответ на ' + str(qs[int(message.from_user.id)]) +
                             ' вопрос')
    state = dp.current_state(user=message.from_user.id)
    await state.set_state(TestStates.all()[int(3)])


@dp.message_handler(state=TestStates.TEST_STATE_3)
async def process_ans_command(message: types.Message):
    w = openpyxl.load_workbook('users.xlsx')
    s = w['users']
    s = w.active
    sh = sh2
    if str(s.cell(row=users[int(message.from_user.id)] + 2, column=5).value) == str(1) and qs[int(message.from_user.id)] <= 20:
        sh = sh1
    elif str(s.cell(row=users[int(message.from_user.id)] + 2, column=5).value) == str(1) and qs[int(message.from_user.id)] > 20:
        cell = s.cell(row=users[int(message.from_user.id)] + 2, column=5)
        cell.value = 2
        w.save('users.xlsx')
        w.close()
        qs[int(message.from_user.id)] = 1

    if (message.text == str(sh.cell(row=qs[int(message.from_user.id)], column=1).value)) or (
            message.text == str(sh.cell(row=qs[int(message.from_user.id)], column=2).value)) or (
            message.text == str(sh.cell(row=qs[int(message.from_user.id)], column=3).value)) or (
            message.text == str(sh.cell(row=qs[int(message.from_user.id)], column=4).value)):
        if (qs[int(message.from_user.id)] <= 10 and str(s.cell(row=users[int(message.from_user.id)] + 2, column=5).value) == str(1)) or (qs[int(message.from_user.id)] <= 5 and str(s.cell(row=qs[int(message.from_user.id)], column=5).value) == str(2)):
            score[int(message.from_user.id)] = int(score[int(message.from_user.id)]) + 1
        else:
            score[int(message.from_user.id)] = int(score[int(message.from_user.id)]) + 2
    state = dp.current_state(user=message.from_user.id)
    if qs[int(message.from_user.id)] == 20 and str(s.cell(row=users[int(message.from_user.id)] + 2, column=5).value) == str(1):
        if int(message.from_user.id) == int(765839138):
            await state.set_state(TestStates.all()[int(0)])
        else:
            await state.set_state(TestStates.all()[int(3)])
        await message.answer('Первый раунд викторины закончен!\n\n'
                             'Второй раунд будет после второго фильма, а пока настало время'
                             'отдыха\n\n!Начать второй раунд можно только по команде от ведущих!'
                             '\n\nЧтобы начать, просто нажми —> /quiz')
        qs[int(message.from_user.id)] += 1
        wb = openpyxl.load_workbook('users.xlsx')
        sheet = wb['users']
        sheet = wb.active
        cell = sheet.cell(row=users[int(message.from_user.id)] + 2, column=3)
        cell.value = score[int(message.from_user.id)]
        wb.save('users.xlsx')
        wb.close()
    elif qs[int(message.from_user.id)] == 10 and str(s.cell(row=users[int(message.from_user.id)] + 2, column=5).value) == str(2):
        await message.answer('Викторина окончена, большое спасибо за участие\n\n'
                             'Результаты будут объявлены чуть позже'
                             '\n\nУзнать свой счет —> /score')
        if int(message.from_user.id) == int(765839138):
            await state.set_state(TestStates.all()[int(0)])
        else:
            await state.set_state(TestStates.all()[int(5)])
        wb = openpyxl.load_workbook('users.xlsx')
        sheet = wb['users']
        sheet = wb.active
        cell = sheet.cell(row=users[int(message.from_user.id)] + 2, column=3)
        cell.value = score[int(message.from_user.id)]
        wb.save('users.xlsx')
        wb.close()
    else:
        qs[int(message.from_user.id)] += 1
        await state.set_state(TestStates.all()[int(2)])
        await process_quiz_command(message)


@dp.message_handler(state=TestStates.TEST_STATE_1 | TestStates.TEST_STATE_0, commands=['user'])
async def process_user_command(message: types.Message):
    wb = openpyxl.load_workbook('users.xlsx')
    sheet = wb['users'].active
    wb.save('users.xlsx')
    wb.close()
    cell = sheet.cell(row=users[int(message.from_user.id)] + 1, column=3)
    cell.value = 0
    state = dp.current_state(user=message.from_user.id)
    await state.set_state(TestStates.all()[int(10)])
    await message.answer('Введи своё настоящее Имя и Фамилию:')


@dp.message_handler(state=TestStates.TEST_STATE_6)
async def process_fio_command(message: types.Message):
    wb = openpyxl.load_workbook('users.xlsx')
    sheet = wb['users']
    sheet = wb.active
    cell = sheet.cell(row=users[int(message.from_user.id)] + 2, column=2)
    cell.value = message.text
    wb.save('users.xlsx')
    wb.close()

    state = dp.current_state(user=message.from_user.id)
    await state.set_state(TestStates.all()[int(7)])
    await message.answer('Прикрепи, пожалуйста, ссылку на ВК, чтобы мы могли связаться с тобой :)')


@dp.message_handler(state=TestStates.TEST_STATE_7)
async def process_link_command(message: types.Message):
    wb = openpyxl.load_workbook('users.xlsx')
    sheet = wb['users']
    sheet = wb.active
    cell = sheet.cell(row=users[int(message.from_user.id)] + 2, column=4)
    cell.value = message.text
    wb.save('users.xlsx')
    wb.close()

    state = dp.current_state(user=message.from_user.id)
    if int(message.from_user.id) == int(765839138):
        await state.set_state(TestStates.all()[int(0)])
    else:
        await state.set_state(TestStates.all()[int(2)])
    await message.answer('Регистрация успешно пройдена!'
                         '\n\nНажми —> /info <— чтобы узнать, что делать дальше')


@dp.message_handler(state=TestStates.TEST_STATE_0, commands=['score'])
async def process_score_command(message: types.Message):
    await message.answer('Твой счет: ' + str(score[int(message.from_user.id)]))


@dp.message_handler(state=TestStates.all())
async def process_no_command(message: types.Message):
    await message.answer(MESSAGES['no_command'])


@dp.message_handler()
async def echo_message(msg: types.Message):
    await bot.send_message(msg.from_user.id, '/start чтобы начать')


async def shutdown(dispatcher: Dispatcher):
    await dispatcher.storage.close()
    await dispatcher.storage.wait_closed()


if __name__ == '__main__':
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    button_1 = types.KeyboardButton(
        text="/start")  # тут мы делаем кнопку старт, если нужно добавить другие или редактировать эту - нужно посмотреть руководство
    keyboard.add(
        button_1)  # https://mastergroosha.github.io/telegram-tutorial-2/buttons/  понятное руководство по кнопкам
    executor.start_polling(dp, on_shutdown=shutdown)
