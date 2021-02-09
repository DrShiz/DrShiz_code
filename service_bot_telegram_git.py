from ldap3 import Server, Connection, SUBTREE, ALL, MODIFY_REPLACE, NTLM
from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from ldap3.extend.microsoft.addMembersToGroups import ad_add_members_to_groups as addUsersInGroups
import datetime
import socket
import pyodbc
import requests
from config import *

class BotStates(StatesGroup):
    """ Возможные состояния бота"""
    START = State()
    SEARCH_USER = State()
    SEARCH_USER_COMF = State()
    CHOOSE_USER = State()
    WAIT_COMMAND = State()
    PWD_PROLONGATE = State()
    PWD_ACCEPT_RESET = State()
    PWD_RESET = State()
    ADD_365 = State()
    UNLOCK_USER = State()
    ACCEPT_ADD_PRINTER = State()
    ADD_ANOTHER_PRINTER = State()
    ADD_PRINTER = State()
    ADD_VPN = State()

# Соединение с DC
server = Server(f'ldaps://{AD_SERVER}:636', use_ssl=True, get_info=ALL)
server_2 = Server(f'ldaps://{AD_SERVER_2}:636', use_ssl=True, get_info=ALL)
server_3 = Server(f'ldaps://{AD_SERVER_3}:636', use_ssl=True, get_info=ALL)
conn = Connection(server, user=AD_USER, password=AD_PASSWORD, authentication=NTLM)
conn.start_tls()
conn_2 = Connection(server_2, user=AD_USER, password=AD_PASSWORD, authentication=NTLM)
conn_2.start_tls()
conn_3 = Connection(server_3, user=AD_USER, password=AD_PASSWORD, authentication=NTLM)
conn_3.start_tls()

# Создание бота
bot = Bot(token)
dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())

# Клавиатура с действиями над пользователем
inline_btn_pwd_prolongate = types.InlineKeyboardButton('Продлить пароль', callback_data='/pwd_prolongate')
inline_btn_pwd_reset = types.InlineKeyboardButton('Сбросить пароль', callback_data='/pwd_reset')
inline_btn_add_365attr = types.InlineKeyboardButton('Добавить атрибут 365', callback_data='/add_365attr')
inline_btn_unlock_user = types.InlineKeyboardButton('Разблокировать пользователя', callback_data='/unlock_user')
inline_btn_add_printer = types.InlineKeyboardButton('Добавить принтер', callback_data='/add_printer')
inline_btn_add_vpn = types.InlineKeyboardButton('Дать доступ VPN', callback_data='/add_vpn')
inline_btn_show_groups = types.InlineKeyboardButton('Вывести список групп пользователя', callback_data='/show_groups')
actions_keyboard = types.InlineKeyboardMarkup(resize_keyboard=True)\
    .add(inline_btn_pwd_prolongate)\
    .add(inline_btn_pwd_reset)\
    .add(inline_btn_add_365attr)\
    .add(inline_btn_unlock_user)\
    .add(inline_btn_add_printer)\
    .add(inline_btn_add_vpn)\
    .add(inline_btn_show_groups)

# Клавиатура с клавишами по-умолчанию
default_keyboard = types.ReplyKeyboardMarkup().add('/search_user').add('/search_user_comf')

# Клавиатура с ответами Да/Нет
inline_btn_yes = types.InlineKeyboardButton('Да', callback_data='/yes')
inline_btn_no = types.InlineKeyboardButton('Нет', callback_data='/no')
answer_keyboard = types.InlineKeyboardMarkup(resize_keyboard=True).row(inline_btn_yes, inline_btn_no)


@dp.message_handler(state='*', commands=['start'])
async def process_start_command(message) -> None:
    """Реакция бота на команду /start"""
    if str(message.chat.id) in Allowed_IDs:
        await bot.send_message(chat_id=message.chat.id,
                               text=f'{IDs[str(message.chat.id)][0]} - доступ разрешен.',
                               reply_markup=default_keyboard)
        await BotStates.START.set()
    else:
        await bot.send_message(chat_id=message.chat.id, text=f'{message.chat.id} - доступ не разрешен.')
        await BotStates.START.set()


@dp.message_handler(state='*', commands=['search_user'])
async def process_search_user_command(message: types.Message) -> None:
    """Реакция бота на команду /search_user"""
    if str(message.chat.id) in Allowed_IDs:
        await bot.send_message(chat_id=message.chat.id, text='Введите данные пользователя (логин/ФИО/e-mail):')
        await BotStates.SEARCH_USER.set()
    else:
        await bot.send_message(chat_id=message.chat.id, text=f'{message.chat.id} - доступ не разрешен.')


@dp.message_handler(state='*', commands=['search_user_comf'])
async def process_search_user_command(message: types.Message) -> None:
    """Реакция бота на команду /search_user_comf"""
    if str(message.chat.id) in Allowed_IDs:
        await bot.send_message(chat_id=message.chat.id, text='Введите данные пользователя (логин/ФИО/e-mail):')
        await BotStates.SEARCH_USER_COMF.set()
    else:
        await bot.send_message(chat_id=message.chat.id, text=f'{message.chat.id} - доступ не разрешен.')


@dp.message_handler(state=BotStates.SEARCH_USER)
async def handle_search_user(message: types.Message):
    """Бот в состоянии поиска пользователей"""
    await search_user_func(message, 1)
    if not len(IDs[str(message.chat.id)][1]['DATA']) <= 1:
        await bot.send_message(chat_id=message.chat.id, text='Укажите номер нужного вам пользователя:')
        await BotStates.CHOOSE_USER.set()
    else:
        await BotStates.WAIT_COMMAND.set()


@dp.message_handler(state=BotStates.SEARCH_USER_COMF)
async def handle_search_user(message: types.Message):
    """Бот в состоянии поиска пользователей"""
    await search_user_func(message, 2)
    if not len(IDs[str(message.chat.id)][1]['DATA']) <= 1:
        await bot.send_message(chat_id=message.chat.id, text='Укажите номер нужного вам пользователя:')
        await BotStates.CHOOSE_USER.set()
    else:
        await BotStates.WAIT_COMMAND.set()


@dp.message_handler(state=BotStates.CHOOSE_USER)
async def handle_choose_user(message: types.Message):
    """Бот в состоянии выбора пользователя"""
    await choose_user_func(message)


async def search_user_func(message, connect_number):
    """Функция поиска пользователей по логину/ФИО/e-mail"""
    if connect_number == 1:
        conn.bind()
        search_user = message.text + '*'
        conn.search(AD_SEARCH_TREE,
                    f'(&(objectCategory=person)(objectClass=user)(|(sAMAccountName={search_user})(cn={search_user})'
                    f'(mail={search_user})))',
                    SUBTREE,
                    attributes=['DistinguishedName', 'sAMAccountName', 'CN', 'pwdLastSet', 'mail',
                                'badPasswordTime', 'badPwdCount', 'lockoutTime', 'lastLogon', 'extensionAttribute4',
                                'ipPhone', 'telephoneNumber', 'mobile', 'objectSid', 'pager', 'userAccountControl',
                                'Company', 'department', 'title', 'memberOf'])

        IDs[str(message.chat.id)][1]['DATA'] = conn.entries
    elif connect_number == 2:
        conn_2.bind()
        search_user = message.text + '*'
        conn_2.search(AD_SEARCH_TREE_2,
                    f'(&(objectCategory=person)(objectClass=user)(|(sAMAccountName={search_user})(cn={search_user})'
                    f'(mail={search_user})))',
                    SUBTREE,
                    attributes=['DistinguishedName', 'sAMAccountName', 'CN', 'pwdLastSet', 'mail',
                                'badPasswordTime', 'badPwdCount', 'lockoutTime', 'lastLogon', 'extensionAttribute4',
                                'ipPhone', 'telephoneNumber', 'mobile', 'objectSid', 'pager', 'userAccountControl',
                                'Company', 'department', 'title', 'memberOf'])

        IDs[str(message.chat.id)][1]['DATA'] = conn_2.entries

    """ Вывод списка пользователей: """
    if len(IDs[str(message.chat.id)][1]['DATA']) == 0:
        """ Если пользователь не найден """
        await bot.send_message(chat_id=message.chat.id, text='Пользователи не найдены')
        await BotStates.START.set()
        conn.unbind()
    elif len(IDs[str(message.chat.id)][1]['DATA']) == 1:
        """Если найден 1 пользователь"""
        # Запоминаем в переменной USERDN DistinguishedName пользователя
        IDs[str(message.chat.id)][1]['USERDN'] = str(IDs[str(message.chat.id)][1]['DATA'][0].DistinguishedName)
        # Запоминаем в переменной NUMBER номер нужного пользователя в списке найденных пользователей
        IDs[str(message.chat.id)][1]['NUMBER'] = 0
        await bot.send_message(chat_id=message.chat.id, text='Данные пользователя:')
        await bot.send_message(chat_id=message.chat.id, text=print_user_info_func(0, message.chat.id),
                               reply_markup=actions_keyboard)
        conn.unbind()
    else:
        """Если найдено несколько пользователей"""
        await bot.send_message(chat_id=message.chat.id, text='Данные пользователя:')
        for entry, i in zip(IDs[str(message.chat.id)][1]['DATA'],
                            range(0, len(IDs[str(message.chat.id)][1]['DATA']))):
            await bot.send_message(chat_id=message.chat.id, text=f'{i+1}, {entry.CN}\n'
                                                                 f'{entry.sAMAccountName}, {entry.mail}\n')


async def choose_user_func(message):
    """Функция выбора пользователя для работы"""
    n = int(message.text)
    # Запоминаем в переменной NUMBER номер нужного пользователя в списке найденных пользователей
    IDs[str(message.chat.id)][1]['NUMBER'] = n - 1
    try:
        await bot.send_message(chat_id=message.chat.id,
                               text=print_user_info_func(n - 1, message.chat.id),
                               reply_markup=actions_keyboard)
        # Запоминаем в переменной USERDN DistinguishedName пользователя
        IDs[str(message.chat.id)][1]['USERDN'] = str(IDs[str(message.chat.id)][1]['DATA'][n - 1].DistinguishedName)
        await BotStates.WAIT_COMMAND.set()
        conn.unbind()
    except Exception as e:
        await bot.send_message(chat_id=message.chat.id, text='Введено неверное значение.')
        await BotStates.CHOOSE_USER.set()


@dp.callback_query_handler(lambda c: c.data == '/pwd_prolongate', state=BotStates.WAIT_COMMAND)
async def handle_wait_command(callback_query: types.CallbackQuery):
    """ Бот в состоянии WAIT_COMMAND - ожидание команды
    Реакция бота на команду /pwd_prolongate - продление пароля """
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Продлить пароль?',
                           reply_markup=answer_keyboard)
    await BotStates.PWD_PROLONGATE.set()


@dp.callback_query_handler(lambda c: c.data == '/pwd_reset', state=BotStates.WAIT_COMMAND)
async def handle_wait_command(callback_query: types.CallbackQuery):
    """ Бот в состоянии WAIT_COMMAND - ожидание команды
    Реакция бота на команду /pwd_reset - сброс пароля """
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Сбросить пароль пользователю?',
                           reply_markup=answer_keyboard)
    await BotStates.PWD_ACCEPT_RESET.set()


@dp.callback_query_handler(lambda c: c.data == '/unlock_user', state=BotStates.WAIT_COMMAND)
async def handle_wait_command(callback_query: types.CallbackQuery):
    """ Бот в состоянии WAIT_COMMAND - ожидание команды
    Реакция бота на команду /unlock_user - разблокировка пользователя"""
    if IDs[str(callback_query.from_user.id)][1]['BLOCKED'] == 1:
        await bot.answer_callback_query(callback_query.id)
        await bot.send_message(callback_query.from_user.id, text='Разблокировать пользователя?',
                               reply_markup=answer_keyboard)
        await BotStates.UNLOCK_USER.set()
    else:
        await bot.answer_callback_query(callback_query.id)
        await bot.send_message(callback_query.from_user.id, text='Пользователь не заблокирован. '
                                                                 'Разблокировка не требуется.')
        await BotStates.WAIT_COMMAND.set()


@dp.callback_query_handler(lambda c: c.data == '/add_365attr', state=BotStates.WAIT_COMMAND)
async def handle_wait_command(callback_query: types.CallbackQuery):
    """Бот в состоянии WAIT_COMMAND - ожидание команды
    Реакция бота на команду /add_365attr - установка атрибута 365"""
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Добавить атрибут 365?',
                           reply_markup=answer_keyboard)
    await BotStates.ADD_365.set()


@dp.callback_query_handler(lambda c: c.data == '/add_vpn', state=BotStates.WAIT_COMMAND)
async def handle_wait_command(callback_query: types.CallbackQuery):
    """Бот в состоянии WAIT_COMMAND - ожидание команды
    Реакция бота на команду /add_vpn - предоставление доступа к VPN"""
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Дать доступ VPN?',
                           reply_markup=answer_keyboard)
    await BotStates.ADD_VPN.set()


@dp.callback_query_handler(lambda c: c.data == '/add_printer', state=BotStates.WAIT_COMMAND)
async def handle_wait_command(callback_query: types.CallbackQuery):
    """ Бот в состоянии WAIT_COMMAND - ожидание команды
    Реакция бота на команду /add_printer - добавление принтера"""
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Добавить принтер?',
                           reply_markup=answer_keyboard)
    await BotStates.ACCEPT_ADD_PRINTER.set()


@dp.message_handler(state=BotStates.PWD_PROLONGATE)
async def handle_pwd_prolongate(message):
    """Бот в состоянии PWD_PROLONGATE - продление пароля"""
    await pwd_prolongate_func(message)


@dp.message_handler(state=BotStates.PWD_ACCEPT_RESET)
async def handle_pwd_accept_reset(message):
    """Бот в состоянии PWD_ACCEPT_RESET - сбрасывать ли пароль"""
    await pwd_accept_reset_func(message)


@dp.message_handler(state=BotStates.PWD_RESET)
async def handle_pwd_reset(message):
    """Бот в состоянии PWD_RESET - сброс пароля"""
    await pwd_reset_func(message, IDs[str(message.chat.id)][1]['USERDN'])


@dp.message_handler(state=BotStates.ADD_365)
async def handle_add_365attr(message):
    """Бот в состоянии ADD_365 - установка атрибута 365"""
    await add_365attr_func(message, IDs[str(message.chat.id)][1]['USERDN'])


@dp.message_handler(state=BotStates.ADD_VPN)
async def handle_add_vpn(message):
    """Бот в состоянии ADD_365 - установка атрибута 365"""
    await add_vpn_func(message, IDs[str(message.chat.id)][1]['USERDN'])


@dp.message_handler(state=BotStates.UNLOCK_USER)
async def handle_unlock_user(message):
    """Бот в состоянии UNLOCK_USER - разблокировка пользователя"""
    await unlock_user_func(message, IDs[str(message.chat.id)][1]['USERDN'])


@dp.message_handler(state=BotStates.ACCEPT_ADD_PRINTER)
async def handle_accept_add_printer(message):
    """Бот в состоянии ACCEPT_ADD_PRINTER - подключать ли принтер"""
    await accept_add_printer_func(message)


@dp.message_handler(state=BotStates.ADD_PRINTER)
async def handle_add_printer(message):
    """Бот в состоянии ADD_PRINTER - добавление принтера"""
    await add_printer_func(message, IDs[str(message.chat.id)][1]['USERDN'])


@dp.callback_query_handler(lambda c: c.data == '/yes', state=BotStates.PWD_PROLONGATE)
async def pwd_prolongate_func(callback_query: types.CallbackQuery):
    """Функция продления пароля"""
    await bot.answer_callback_query(callback_query.id)
    try:
        conn.bind()
        conn.modify(IDs[str(callback_query.from_user.id)][1]['USERDN'], {'pwdLastSet': [(MODIFY_REPLACE, 0)]})
        conn.modify(IDs[str(callback_query.from_user.id)][1]['USERDN'], {'pwdLastSet': [(MODIFY_REPLACE, -1)]})
        await bot.send_message(callback_query.from_user.id, text='Пароль успешно продлен.')
        await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
        await BotStates.WAIT_COMMAND.set()
        conn.unbind()
    except Exception as e:
        await bot.send_message(callback_query.from_user.id, text='Пароль не удалось продлить.')
        await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
        await BotStates.WAIT_COMMAND.set()


@dp.callback_query_handler(lambda c: c.data == '/no', state='*')
async def cancel_operation(callback_query: types.CallbackQuery):
    """ Реакция бота на ответ 'Нет'"""
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Операция отменена.')
    await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
    await BotStates.WAIT_COMMAND.set()


@dp.callback_query_handler(lambda c: c.data == '/yes', state=BotStates.PWD_ACCEPT_RESET)
async def pwd_accept_reset_func(callback_query: types.CallbackQuery):
    """Функция подтверждения сброса пароля"""
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Введите новый пароль.')
    await BotStates.PWD_RESET.set()


async def pwd_reset_func(message, userdn):
    """Функция сброса пароля"""
    userpswd = str(message.text)
    if pwd_check_func(userpswd):
        try:
            conn.bind()
            conn.extend.microsoft.modify_password(userdn, userpswd)
            await bot.send_message(chat_id=message.chat.id, text='Пароль сменен.')
            await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
            await BotStates.WAIT_COMMAND.set()
            conn.unbind()
        except Exception as e:
            await bot.send_message(chat_id=message.chat.id, text='Сменить пароль не удалось.')
            await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
            await BotStates.WAIT_COMMAND.set()
    else:
        await bot.send_message(chat_id=message.chat.id, text='Пароль не соотвествует требованиям:\n'
                                                             'Минимальная длина - 8 символов\n'
                                                             'Выполняется 3 из 4 условий:\n'
                                                             'Наличие большой буквы\n'
                                                             'Наличие маленькой буквы\n'
                                                             'Наличие цифры\n'
                                                             'Наличие спецсимвола\n\n'
                                                             'Введите другой пароль:')
        await BotStates.PWD_RESET.set()


@dp.callback_query_handler(lambda c: c.data == '/yes', state=BotStates.ADD_365)
async def add_365attr_func(callback_query: types.CallbackQuery):
    """Добавление extensionAttribute4: 365 пользователю"""
    try:
        conn.bind()
        conn.modify(IDs[str(callback_query.from_user.id)][1]['USERDN'], {'extensionAttribute4': [(MODIFY_REPLACE, 365)]})
        await bot.answer_callback_query(callback_query.id)
        await bot.send_message(callback_query.from_user.id, text='Атрибут 365 добавлен.')
        await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
        await BotStates.WAIT_COMMAND.set()
        conn.unbind()
    except Exception as e:
        await bot.answer_callback_query(callback_query.id)
        await bot.send_message(callback_query.from_user.id, text='Не удалось добавить атрибут.')
        await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
        await BotStates.WAIT_COMMAND.set()


@dp.callback_query_handler(lambda c: c.data == '/yes', state=BotStates.UNLOCK_USER)
async def unlock_user_func(callback_query: types.CallbackQuery):
    """Функция разблокировки пользователя"""
    try:
        conn.bind()
        conn.modify(IDs[str(callback_query.from_user.id)][1]['USERDN'], {'lockoutTime': [(MODIFY_REPLACE, 0)]})
        await bot.answer_callback_query(callback_query.id)
        await bot.send_message(callback_query.from_user.id, text='Пользователь разблокирован.')
        await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
        await BotStates.WAIT_COMMAND.set()
        conn.unbind()
    except Exception as e:
        await bot.answer_callback_query(callback_query.id)
        await bot.send_message(callback_query.from_user.id, text='Не удалось разблокировать пользователя.')
        await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
        await BotStates.WAIT_COMMAND.set()


# Словарь статусов УЗ пользователя
USER_STATUS = {'2': 'Учетная запись отключена',
               '5': 'Учетная запись заблокирована',
               '10': 'Учетная запись по умолчанию. Обычная активная учетная запись',
               '17': 'Срок действия пароля не ограничен',
               '24': 'Срок действия пароля пользователя истек'}


def check_user_status_func(uac, lockouttime, message_chat_id):
    """Функция получения статуса УЗ пользователя"""
    status = []
    uac_str = str(bin(int(str(uac))))
    for i in range(1, len(uac_str)-1):
        if uac_str[-i] == '1':
            try:
                status.append(USER_STATUS[str(i)])
            except Exception as e:
                pass
    if lockouttime and str(lockouttime) != '1601-01-01 00:00:00+00:00':
        status.append('Учетная запись заблокирована')
        IDs[str(message_chat_id)][1]['BLOCKED'] = 1
    else:
        IDs[str(message_chat_id)][1]['BLOCKED'] = 0
    return status


def print_user_info_func(n, message_chat_id) -> str:
    """Функция, печатающая инфо о пользователе"""
    data = IDs[str(message_chat_id)][1]['DATA']
    try:
        """Попытка получить ip-адрес ПК по его имени"""
        ip = socket.gethostbyname(str(data[n].pager))
    except Exception as e:
    	ip = []
    text = f'{data[n].CN}\n'\
           f'{data[n].sAMAccountName}, {data[n].mail}\n' \
           f'{data[n].DistinguishedName}\n'\
           f'{check_attribute_func("Профиль - ", get_user_url_from_home(data[n].sAMAccountName))}'\
           f'{check_attribute_func("Телеграм - ", get_telegram_from_home(data[n].mail))}'\
           f'{check_attribute_func("Организация - ", data[n].Company)}' \
           f'{check_attribute_func("Департамент - ", data[n].department)}'\
           f'{check_attribute_func("Должность - ", data[n].title)}'\
           'Статус пользователя - '\
           f'{check_user_status_func(data[n].userAccountControl, data[n].lockoutTime, message_chat_id)}\n'\
           'Дата установки пароля - '\
           f'{(data[n].pwdLastSet.values[0].strftime("%m/%d/%Y, %H:%M:%S, %z"))}\n' \
           f'{pwd_age_func(data[n].pwdLastSet.values[0])}'\
           f'{check_attribute_func("Имя ПК - ", data[n].pager)}' \
           f'{check_attribute_func("IP-адрес ПК - ", ip)}'\
           f'{check_attribute_func("AnyDesk - ", full_pc_name_to_short(data[n].pager.values))}' \
           f'{check_attribute_func("Teamviewer_ID - ", get_teamviewer_id(str(data[n].pager)[:9]))}' \
           f'{get_mac_name_func(data[n].sAMAccountName, jamf_userAndPass)}'\
           f'{check_attribute_func("IP-телефон - ", data[n].ipPhone)}' \
           f'{check_attribute_func("Телефон - ", data[n].telephoneNumber)}' \
           f'{check_attribute_func("Мобильный - ", data[n].mobile)}' \
           'extensionAttribute4 - ' \
           f'{data[n].extensionAttribute4}\n\n' \
           'Что сделать с пользователем?'
    return text


def pwd_check_func(passwd):
    """Функция проверки удовлетворения требованиям пароля"""
    specialsym = ['%', '*', ')', '?', '@', '#', '$', '~', '-']
    i = 0     # Счетчик условий пригодности пароля
    if any(char.isdigit() for char in passwd):
        i += 1
    if any(char.isupper() for char in passwd):
        i += 1
    if any(char.islower() for char in passwd):
        i += 1
    if any(char in specialsym for char in passwd):
        i += 1
    if len(passwd) >= 8 and i >= 3:
        return True


@dp.callback_query_handler(lambda c: c.data == '/yes', state=BotStates.ACCEPT_ADD_PRINTER)
async def accept_add_printer_func(callback_query: types.CallbackQuery):
    """Функция подтверждения добавления принтера"""
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, text='Введите номер принтера.')
    await BotStates.ADD_PRINTER.set()

async def add_printer_func(message, userdn):
    """Функция добавления в принтерную группу"""
    if str(IDs[str(message.chat.id)][1]['DATA'][0].DistinguishedName)[-26:] != 'DC=comf,DC=PICompany,DC=ru':
        conn.bind()
        printer = message.text
        prngroup = 'prn-prn01pik-'+str(printer)+'-pr-g'
        conn.search(AD_SEARCH_TREE, '(&(objectCategory=Group)'f'(cn={prngroup}))', SUBTREE,
                    attributes=['DistinguishedName'])
        try:
            groupdn = str(conn.entries[0].DistinguishedName)
        except Exception as e:
            await bot.send_message(chat_id=message.chat.id, text='Принтер не найден')
        groups_list_full = IDs[str(message.chat.id)][1]['DATA'] \
            [IDs[str(message.chat.id)][1]['NUMBER']].memberOf.values
        groups_list = []
        for group in groups_list_full:
            groups_list.append(group.rsplit(',', -1)[0][3:])
        if prngroup not in groups_list:
            try:
                addUsersInGroups(conn, userdn, groupdn)
                await bot.send_message(chat_id=message.chat.id, text='В группу добавлен')
                await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
                await BotStates.WAIT_COMMAND.set()
                conn.unbind()
            except Exception as e:
                await bot.send_message(chat_id=message.chat.id, text='Добавить в группу не удалось')
                await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
                await BotStates.WAIT_COMMAND.set()
                conn_3.unbind()
        else:
            await bot.send_message(chat_id=message.chat.id, text='Пользователь уже в группе')
            await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
            await BotStates.WAIT_COMMAND.set()
            conn_3.unbind()
    else:
        conn_3.bind()
        printer = message.text
        prngroup = 'prn-prn01pik-'+str(printer)+'-pr-l'
        conn_3.search(AD_SEARCH_TREE_3, '(&(objectCategory=Group)'f'(cn={prngroup}))', SUBTREE,
                    attributes=['DistinguishedName'])
        try:
            groupdn = str(conn_3.entries[0].DistinguishedName)
            print(userdn)
            print(groupdn)
        except Exception as e:
            await bot.send_message(chat_id=message.chat.id, text='Принтер не найден')
        groups_list_full = IDs[str(message.chat.id)][1]['DATA'] \
            [IDs[str(message.chat.id)][1]['NUMBER']].memberOf.values
        groups_list = []
        for group in groups_list_full:
            groups_list.append(group.rsplit(',', -1)[0][3:])
        if prngroup not in groups_list:
            try:
                addUsersInGroups(conn_3, userdn, groupdn)
                await bot.send_message(chat_id=message.chat.id, text='В группу добавлен')
                await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
                await BotStates.WAIT_COMMAND.set()
                conn_3.unbind()
            except Exception as e:
                print(e)
                await bot.send_message(chat_id=message.chat.id, text='Добавить в группу не удалось')
                await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
                await BotStates.WAIT_COMMAND.set()
                conn_3.unbind()
        else:
            await bot.send_message(chat_id=message.chat.id, text='Пользователь уже в группе')
            await bot.send_message(chat_id=message.chat.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
            await BotStates.WAIT_COMMAND.set()
            conn_3.unbind()


@dp.callback_query_handler(lambda c: c.data == '/yes', state=BotStates.ADD_VPN)
async def add_vpn_func(callback_query: types.CallbackQuery):
    """Добавление в группу с VPN доступом"""
    conn_3.bind()
    vpngroup = 'adm-vpn-crt_users_allow-u'
    conn_3.search(AD_SEARCH_TREE_3, '(&(objectCategory=Group)'f'(cn={vpngroup}))', SUBTREE,
                  attributes=['DistinguishedName'])
    groupdn = str(conn_3.entries[0].DistinguishedName)
    groups_list_full = IDs[str(callback_query.from_user.id)][1]['DATA'] \
        [IDs[str(callback_query.from_user.id)][1]['NUMBER']].memberOf.values
    groups_list = []
    for group in groups_list_full:
        groups_list.append(group.rsplit(',', -1)[0][3:])
    if vpngroup not in groups_list:
        try:
            addUsersInGroups(conn_3, IDs[str(callback_query.from_user.id)][1]['USERDN'], groupdn)
            await bot.answer_callback_query(callback_query.id)
            await bot.send_message(chat_id=callback_query.from_user.id, text='В группу добавлен')
            await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
            await BotStates.WAIT_COMMAND.set()
            conn_3.unbind()
        except Exception as e:
            print(e)
            await bot.answer_callback_query(callback_query.id)
            await bot.send_message(chat_id=callback_query.from_user.id, text='Добавить в группу не удалось')
            await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
            await BotStates.WAIT_COMMAND.set()
            conn_3.unbind()
    else:
        await bot.answer_callback_query(callback_query.id)
        await bot.send_message(chat_id=callback_query.from_user.id, text='Пользователь уже в группе')
        await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
        await BotStates.WAIT_COMMAND.set()
        conn_3.unbind()


@dp.callback_query_handler(lambda c: c.data == '/show_groups', state=BotStates.WAIT_COMMAND)
async def show_user_groups(callback_query: types.CallbackQuery):
    """Функция вывода групп пользователя"""
    await bot.answer_callback_query(callback_query.id)
    groups_list_full = IDs[str(callback_query.from_user.id)][1]['DATA']\
        [IDs[str(callback_query.from_user.id)][1]['NUMBER']].memberOf.values
    groups_list = []
    for group in groups_list_full:
        groups_list.append(group.rsplit(',', -1)[0][3:])
    for group in groups_list:
        await bot.send_message(callback_query.from_user.id, text=group)
    await bot.send_message(chat_id=callback_query.from_user.id, text='Что сделать с пользователем?',
                               reply_markup=actions_keyboard)
    await BotStates.WAIT_COMMAND.set()


def check_attribute_func(name, attribute):
    """Функция проверки, что атрибут имеет значение"""
    if attribute:
        return f'{name + str(attribute)}\n'
    else:
        return ''


def full_pc_name_to_short(name):
	'''Функция сокращения имени ПК'''
	try:
		return name[0].split('.')[0] + ", password: " + anydesk_pwd
	except:
		return ''


def get_mac_name_func(user, userAndPass):
    """Функция получения имени макбука по CN пользователя из JAMF"""
    headers = {'Authorization': 'Basic %s' % userAndPass, 'Accept': 'application/json'}
    try:
        resp = requests.get(f'{jamf_url}{user}', headers=headers)
        macs_json = resp.json()['user']['links']['computers']
        macs = ''
        for mac in macs_json:
            macs += (mac['name']) + ', '
        macs = macs[:-2]
        return f'{"Имя MAC - " + macs}\n'
    except:
        return ''


def pwd_age_func(date):
    """Вычисление сколько дней осталось у пароля"""
    curr_date = datetime.datetime.now()
    life = 70 - ((curr_date.year - date.year)*365 + (curr_date.month - date.month)*30 + (curr_date.day - date.day))
    if life > 0:
        return f'{"Пароль истечет через - " + str(life) + " дней"}\n'
    else:
        return 'Пароль истек\n'


def get_teamviewer_id(pc_name):
    """Получение Teamviewer ID по имени ПК из SCCM"""
    try:
        with pyodbc.connect(
                'DRIVER=' + driver + ';SERVER=' + sql_server + ';DATABASE=' + sql_database + ';UID=' + sql_username + ';PWD=' + sql_password + ';Trusted_Connection=yes;') as conn:
            with conn.cursor() as cursor:
                cursor.execute(f"SELECT [CM_CM2].[dbo].[TeamViewer_DATA].[ClientID00] ,[CM_CM2].[dbo].[System_DATA].[Name0]\
                               FROM [CM_CM2].[dbo].[TeamViewer_DATA], [CM_CM2].[dbo].[System_DATA]\
                               where [CM_CM2].[dbo].[TeamViewer_DATA].[MachineID] = [CM_CM2].[dbo].[System_DATA].[MachineID]\
                               and [dbo].[System_DATA].[Name0] ='{pc_name}'")
                row = cursor.fetchone()
                try:
                    tv_id = str(row[0]) + ', password: ' + teamviewer_pwd
                    return tv_id
                except:
                    return ''
    except:
        return ''


def get_telegram_from_home(user_email):
    """Получение телеграма пользователя с портала home.pik.ru по его e-mail"""
    user_email = str(user_email).replace('@', '%40')
    headers = {'Authorization': f'Bearer {token_api}', 'Accept': 'application/json'}
    try:
        responce = requests.get(f'{home_api_url}api/v1/Employee/byEmails?emails={user_email}',
                                headers=headers)
        telegram = 'https://t.me/' + responce.json()['data'][0]['telegram']
    except:
        telegram = ''
    return telegram


def get_user_url_from_home(username):
    """Получение ссылки на профиль пользователя с потрала home.pik.ru по его e-mail"""
    username = str(username)
    headers = {'Authorization': f'Bearer {token_api}', 'Accept': 'application/json'}
    ids = []
    urls = []
    try:
        responce = requests.get(f'{home_api_url}api/v1/Employee/byLogin?login=main%5C{username}',
                                headers=headers)
        for resp in responce.json():
            ids.append(resp['id'])
        for id in ids:
            urls.append(f'{home_url}employees/{id}')
        return urls
    except:
        return ''


if __name__ == '__main__':
    """Запуск бота"""
    executor.start_polling(dp)
