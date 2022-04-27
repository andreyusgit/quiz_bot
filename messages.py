# Файл со всеми командами, если нужно добавить команду, то лучше добавить её сюда

from utils import TestStates

help_message = 'Я написан чтобы помочь тебе классно провести время. В определенное ' \
               'время викторина начнется сама и тебе не нужно ничего для этого делать. ' \
               'Пока ждешь - можешь ввести команду /thanks и поблагодарить меня или же ' \
               'с помощью команды /fan ты можешь с интересом подождать начала'


start_message = 'Привет! Тебя приветствует бот-викторина посвященный ночи кино.' \
                '\nНажми -> /help <- чтобы узнать подробнее что я умею'

invalid_key_message = 'Пароль "{key}" не подходит.\n'

state_change_success_message = 'Введён корректный пароль\nТвой институт - "{key}" \nПросто введи Фамилию ' \
                               'Имя Отчество через пробелы и я попробую его найти \n' \
                               'чтобы сменить институт выполните сброс пароля /password'
state_reset_message = 'Пароль успешно сброшен'
current_state_message = 'Текущий институт - "{current_state}""'
thanks = "Спасибо, что воспользовались ботом!\nЛучшей поддержкой будет подписка\nна мой инстаграм - andrey_us_\n" \
         "или же телеграмм канал - https://t.me/andreychanel"
dont_know_command = "К сожалению, я не пока не знаю такую команду"

MESSAGES = {
    'start': start_message,
    'help': help_message,
    'invalid_key': invalid_key_message,
    'state_change': state_change_success_message,
    'state_reset': state_reset_message,
    'current_state': current_state_message,
    'thx': thanks,
    'no_command': dont_know_command
}
