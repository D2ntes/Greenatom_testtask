# -*- coding: utf-8 -*-
"""
Открыть yandex.ru
Нажать на USD MOEX
Когда откроется динамика курса, скопировать в эксель курс за последние 10 дней
В экселе шапка (A) Дата, (B) Курс, (C)Изменение
Повторить для Евро, записать в ячейки (D), (E), (F) дату евро, курс, изменение
Для каждой строки полученного файла поделить курс евро на доллар и полученное значение записать в ячейку (G)
Выровнять – автоширина
Формат чисел – финансовый
Проверить, чтоб автосумма в экселе опознавала ячейки как числовой формат.
Послать итоговый файл отчета себе на почту.
В письме указать количество строк в экселе в правильном склонении
"""

from get_course_excel import get_quote_data, get_link_currency, save_to_excel, format_excel
from send_mail_file import send_email
from os import path
import datetime
from num2str import num2text


def main():
    URL_YANDEX = 'https://yandex.ru/'
    OUTPUT_FILE_NAME = path.dirname(path.realpath(__file__)) + '\\outputdata' + datetime.datetime.today().strftime(
        "_%d%m%Y_%H%M%S") + '.xlsx'

    # Получить данные с сайта и сохранить в файл xls
    for currency in ['USD', 'EUR']:
        data_from_site = get_quote_data(get_link_currency(URL_YANDEX, currency))
        save_to_excel(OUTPUT_FILE_NAME, data_from_site)

    # Форматировать по заданию файл xls и получить кол-во строк в файле
    max_row = format_excel(OUTPUT_FILE_NAME)

    #  Указать количество строк в экселе в правильном склонении
    units = ((u'строка', u'строки', u'строк'), 'f')
    num_rows = num2text(max_row, units)

    # Отправить письмо на почту
    send_email('gruzdev-n@mail.ru', 'Test Greenatom', num_rows, OUTPUT_FILE_NAME)


if __name__ == '__main__':
    main()
