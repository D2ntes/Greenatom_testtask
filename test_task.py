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
import Excel
import Parser
from Action import Context
from os import path
import datetime
from num2str import num2text

URL_YANDEX = 'https://yandex.ru/'
OUTPUT_FILE_NAME = path.dirname(path.realpath(__file__)) + '\\outputdata' + datetime.datetime.today().strftime(
    "_%d%m%Y_%H%M%S") + '.xlsx'


def create_handlers_chain():
    head = Parser.ParseQuotePageAction()
    writer = head.set_next(Excel.WriteToXlsAction(OUTPUT_FILE_NAME))
    return head


def main():
    handlers_chain = create_handlers_chain()

    for col_shift, currency in enumerate(['USD', 'EUR']):
        context = Context()
        context.currency = currency
        context.url = URL_YANDEX
        handlers_chain.handle(context)

    #  Указать количество строк в экселе в правильном склонении
    # units = ((u'строка', u'строки', u'строк'), 'f')
    # num_rows = num2text(rows_count, units)

    # Отправить письмо на почту
    # send_email('gruzdev-n@mail.ru', 'Test Greenatom', num_rows, OUTPUT_FILE_NAME)


if __name__ == '__main__':
    main()
