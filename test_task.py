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

from get_course_excel import parse_quote_data, parse_currency_link, get_page_soup
from excel import apply_sheet_format, ExcelWriter
from send_mail_file import send_email
from os import path
import datetime
from num2str import num2text

URL_YANDEX = 'https://yandex.ru/'
OUTPUT_FILE_NAME = path.dirname(path.realpath(__file__)) + '\\outputdata' + datetime.datetime.today().strftime(
    "_%d%m%Y_%H%M%S") + '.xlsx'


def main():
    writer = ExcelWriter(OUTPUT_FILE_NAME)

    # Получить данные с сайта и сохранить в файл xls
    for shift, currency in enumerate(['USD', 'EUR']):
        headers, rows = process_yandex_quote_page(URL_YANDEX, currency)
        shift = shift * len(headers)
        writer.write_row(row=1, column=1 + shift, row_data=headers)
        writer.write_rows(row=2, column=1 + shift, rows_data=rows)

    rows_count = writer.worksheet.max_row
    writer.save_and_close()

    #  Указать количество строк в экселе в правильном склонении
    units = ((u'строка', u'строки', u'строк'), 'f')
    num_rows = num2text(rows_count, units)

    # Отправить письмо на почту
    # send_email('gruzdev-n@mail.ru', 'Test Greenatom', num_rows, OUTPUT_FILE_NAME)


def process_yandex_quote_page(url, currency):
    yandex_page_soup = get_page_soup(url)
    currency_url_link = parse_currency_link(yandex_page_soup, currency)
    currency_page_soup = get_page_soup(currency_url_link)
    headers, rows = parse_quote_data(currency_page_soup)
    return headers, rows


if __name__ == '__main__':
    main()
