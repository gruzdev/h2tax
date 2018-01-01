#!/usr/bin/env python
# -*- coding: utf-8 -*-

# This is free and unencumbered software released into the public domain.
#
# Anyone is free to copy, modify, publish, use, compile, sell, or
# distribute this software, either in source code form or as a compiled
# binary, for any purpose, commercial or non-commercial, and by any
# means.
#
# In jurisdictions that recognize copyright laws, the author or authors
# of this software dedicate any and all copyright interest in the
# software to the public domain. We make this dedication for the benefit
# of the public at large and to the detriment of our heirs and
# successors. We intend this dedication to be an overt act of
# relinquishment in perpetuity of all present and future rights to this
# software under copyright law.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
# OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
# ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
# OTHER DEALINGS IN THE SOFTWARE.
#
# For more information, please refer to <http://unlicense.org/>

from __future__ import print_function
import autoit
from autoit import properties
from datetime import datetime
from time import sleep
import sys
import ctypes
import os

# Налоговый период.
YEAR = 2017
# Исполняемый файл программы Декларация.
DECL_BINARY = "Decl{}.exe".format(YEAR)
# Полный путь к программе.
PATH = u"{}\АО ГНИВЦ\Декларация {}\{}".format(os.environ['PROGRAMFILES'], YEAR, DECL_BINARY)
# Заголовок основного окна программы.
TITLE_MAIN = u"Без имени - Декларация {}".format(YEAR)
# Координаты кнопки "Доходы за пределами РФ".
X_XRUS, Y_XRUS = 98, 414
# Координаты кнопки "Добавить источник выплат".
X_ADD_SRC, Y_ADD_SRC = 206, 113
# Источник выплаты дохода по умолчанию.
DEFAULT_INCOME_SOURCE = u"Брокер \"XXX\" (США)"
# Код страны по умолчанию (США).
DEFAULT_COUNTRY_CODE = 840
# Код валюты по умолчанию (доллар США).
DEFAULT_CURRENCY_CODE = 840
# Код дохода "Дивиденды".
INCOME_TYPE_DIVIDENDS = 1010


def main():
    # Подготовка к вводу данных:
    # запускаю программу,
    # отмечаю "Имеются доходы - В иностранной валюте",
    # перехожу в раздел "Доходы за пределами РФ",
    # закрываю окно с сообщением "Не введен номер инспекции".
    open_decl()

    # Ввод данных о дивидендах:
    # дата, поступившая сумма, удержанная брокером сумма, описание бумаги.
    # Дата в формате YYYY-MM-DD т.е. ГГГГ-ММ-ДД т.е. ГОД-МЕСЯЦ-ДЕНЬ .
    desc = "ISHARES CORE US AGGREGATE BOND ETF"
    fill_div("2017-02-01", 21.8, 2.18, desc)
    # ...
    fill_div("2017-12-21", 3.64, 0.36, desc)

    desc = "VANGUARD TOTAL BOND MARKET ETF"
    fill_div("2017-02-07", 16.8, 1.68, desc)
    # ...
    fill_div("2017-12-29", 3.7, 0, desc)     # LT Cap Gain
    fill_div("2017-12-29", 17.42, 1.74, desc)

    desc = "POLYMETAL INTERNATIONAL PLC"
    source = u"Брокер \"YYY\" (Россия)"
    country = 832
    fill_div("2017-01-13", 15, 0, desc, source, country)
    fill_div("2017-05-30", 18, 0, desc, source, country)
    # ...

    # Вывожу итог по году.
    stats_print()


def fill_div(date, gross, withheld, sec_desc,
             source=DEFAULT_INCOME_SOURCE,
             country=DEFAULT_COUNTRY_CODE,
             currency=DEFAULT_CURRENCY_CODE
             ):
    # Проверяю, что строка с описанием источника не пустая.
    if source == "":
        sys.exit(u"Не задано описание источника выплаты")
    # Проверяю, что доход - положительный.
    if gross <= 0:
        sys.exit(u"Отрицательный или нулевой доход: {}".format(gross))
    # Проверяю, что удержанная сумма - неотрицательная.
    if withheld < 0:
        sys.exit(u"Отрицательная удержанная сумма: {}".format(withheld))
    # Проверяю, что удержано меньше, чем поступило.
    if withheld >= gross:
        sys.exit(u"Удержано больше(равно), чем поступило: {}".format(withheld))
    # Проверяю, что дата - существующая.
    try:
        dt = datetime.strptime(date, "%Y-%m-%d")
    except ValueError:
        sys.exit(u"Несуществующая дата: {}".format(date))
    # Проверяю, что год введённой даты равен году налогового периода.
    if dt.year != YEAR:
        sys.exit(u"Неверный год в дате: {}".format(date))
    # Проверяю, что дата - рабочий день.
    if dt.isoweekday() > 5:
        sys.exit(u"Дата - выходной: {}".format(date))
    # Нажимаю "Добавить источник выплат".
    autoit.mouse_click("left", X_ADD_SRC, Y_ADD_SRC)
    # Жду появления окна "Доход".
    title_income = u"Доход"
    autoit.win_wait_active(title_income)
    # Заполняю "Наименование источника выплаты"
    income_src = u"{} Дивиденд от {} {}".format(source, dt.strftime("%d.%m.%Y"), sec_desc)
    autoit.control_set_text(title_income, "[CLASS:TMaskedEdit; INSTANCE:2]", income_src)
    # Заполняю "Код страны".
    try:
        open_listview_then_find_and_select_number(title_income, "[CLASS:TButton; INSTANCE:1]", u"Справочник стран", country)
    except ValueError:
        sys.exit(u"Неизвестный код страны: {}".format(country))
    # Нажимаю "Да".
    autoit.control_click(title_income, "[CLASS:TButton; INSTANCE:2]")
    # Жду появления основного окна.
    autoit.win_wait_active(TITLE_MAIN)
    # Заполняю "Дата получения дохода".
    set_datetime_picker(TITLE_MAIN, "[CLASS:TDateTimePicker; INSTANCE:2]", dt.day, dt.month, dt.year)
    # Заполняю "Дата уплаты налога".
    set_datetime_picker(TITLE_MAIN, "[CLASS:TDateTimePicker; INSTANCE:1]", dt.day, dt.month, dt.year)
    # Заполняю "Код и наименование валюты".
    try:
        open_listview_then_find_and_select_number(TITLE_MAIN, "[CLASS:TButton; INSTANCE:3]", u"Справочник валют", currency)
    except ValueError:
        sys.exit(u"Неизвестный код валюты: {}".format(currency))
    # Отмечаю "Автоматическое определение курса валюты".
    autoit.control_click(TITLE_MAIN, "[CLASS:TCheckBox; INSTANCE:1]")
    # Заполняю "Код дохода".
    try:
        open_listview_then_find_and_select_number(TITLE_MAIN, "[CLASS:TButton; INSTANCE:2]", u"Справочник видов доходов", INCOME_TYPE_DIVIDENDS)
    except ValueError:
        sys.exit(u"Неизвестный код дохода: {}".format(INCOME_TYPE_DIVIDENDS))
    # Заполняю "Полученный доход" - "В иностранной валюте"
    autoit.control_set_text(TITLE_MAIN, "[CLASS:TMaskedEdit; INSTANCE:3]", float_to_string_rus(gross))
    # Заполняю "Налог, уплаченный в иностранном государстве" - "В иностранной валюте".
    autoit.control_set_text(TITLE_MAIN, "[CLASS:TMaskedEdit; INSTANCE:2]", float_to_string_rus(withheld))
    # Сохраняю данные для подсчёта итога по году.
    stats_update(source, currency, gross, withheld)


def stats_update(source, currency, gross, withheld):
    if not hasattr(stats_update, "stats"):
        stats_update.stats = {}
    s = stats_update.stats
    if source not in s:
        s[source] = {}
    if currency not in s[source]:
        s[source][currency] = {}
    if 'gross' not in s[source][currency]:
        s[source][currency]['gross'] = 0
    if 'withheld' not in s[source][currency]:
        s[source][currency]['withheld'] = 0
    s[source][currency]['gross'] += gross
    s[source][currency]['withheld'] += withheld
    stats_update.stats = s


def stats_print():
    s = stats_update.stats
    print(u"Итог", YEAR)
    for source in s:
        print(u"  Источник:", source)
        for currency in s[source]:
            print(u"    Код валюты:", currency)
            print(u"      Доход:", s[source][currency]["gross"])
            print(u"      Удержано:", s[source][currency]["withheld"])


def open_listview_then_find_and_select_number(title_parent, ctrl_button_open, title_lv, number_to_select):
    # Нажимаю кнопку вызова списка.
    autoit.control_click(title_parent, ctrl_button_open)
    # Жду появления окна списка.
    autoit.win_wait_active(title_lv)
    # Ищу строку в списке.
    ctrl_lv = "[CLASS:TListView; INSTANCE:1]"
    index = autoit.control_list_view(title_lv, ctrl_lv, "FindItem", extras1=str(number_to_select))
    if index == "-1":
        raise ValueError
    # Выбираю строку в списке.
    autoit.control_list_view(title_lv, ctrl_lv, "Select", extras1=index)
    # Нажимаю кнопку подтвержения выбора.
    ctrl_button_confirm = "[CLASS:TButton; INSTANCE:2]"
    autoit.control_click(title_lv, ctrl_button_confirm)
    # Жду появления родительского окна.
    autoit.win_wait_active(title_parent)


def set_datetime_picker(title, ctrl, day, month, year):
    # Устанавливаю фокус на элемент выбора даты.
    autoit.control_focus(title, ctrl)
    delay_s = 0.1
    # Если дата в элементе вводится не первый раз, то фокус будет на годе.
    # Нажимаю два раза стрелку влево, чтобы установить фокус на день.
    set_datetime_picker.calls += 1
    if set_datetime_picker.calls > 2:
        autoit.control_send(title, ctrl, "{LEFT}")
        sleep(delay_s)
        autoit.control_send(title, ctrl, "{LEFT}")
        sleep(delay_s)
    # Ввожу компонент даты, нажимаю стрелку вправо.
    autoit.control_send(title, ctrl, str(day))
    sleep(delay_s)
    autoit.control_send(title, ctrl, "{RIGHT}")
    sleep(delay_s)
    autoit.control_send(title, ctrl, str(month))
    sleep(delay_s)
    autoit.control_send(title, ctrl, "{RIGHT}")
    sleep(delay_s)
    autoit.control_send(title, ctrl, str(year))
    sleep(delay_s)
set_datetime_picker.calls = 0


def open_decl():
    # Проверяю, что разрешение экрана не менее 1024 в ширину.
    user32 = ctypes.windll.user32
    resx, resy = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    if resx < 1024:
        sys.exit(u"Разрешение экрана не поддерживается: {}x{}".format(resx, resy))
    # Проверяю, что Декларация не запущена.
    if autoit.process_exists(DECL_BINARY):
        sys.exit(u"Программа Декларация уже запущена")
    # Запускаю программу.
    autoit.run(PATH)
    # Жду появления основного окна.
    autoit.win_wait_active(TITLE_MAIN)
    # Разворачиваю окно на весь экран.
    autoit.win_set_state(TITLE_MAIN, properties.SW_MAXIMIZE)
    # Отмечаю "Имеются доходы - В иностранной валюте".
    autoit.control_click(TITLE_MAIN, u"[TEXT:В иностранной валюте]")
    # Перехожу в раздел "Доходы за пределами РФ".
    autoit.mouse_click("left", X_XRUS, Y_XRUS)
    # Жду появления окна с сообщением "Не введен номер инспекции".
    title_err = u"Проверка: Задание условий"
    autoit.win_wait_active(title_err)
    # Нажимаю "Пропустить".
    autoit.control_click(title_err, u"[TEXT:Пропустить]")
    # Жду появления основного окна.
    autoit.win_wait_active(TITLE_MAIN)


def float_to_string_rus(val):
    # Преобразую число в строку, заменяю точку на запятую.
    return str(val).replace(".", ",")


if __name__ == '__main__':
    main()
