# -*- coding: utf-8 -*-
from datetime import datetime
from time import sleep

import requests
from bs4 import BeautifulSoup

from Action import Action


class ParseQuotePageAction(Action):
    def __init__(self):
        self.data_types = [
            lambda x: datetime.strptime(f"{x}", "%d.%m.%y").date(),
            lambda x: float(x.replace(',', '.')),
            lambda x: float(x.replace(',', '.')),
        ]

    def process(self, context):
        page = self.get_page_soup(context.url)
        currency_url = self.parse_currency_link(page, context.currency)
        page = self.get_page_soup(currency_url)
        headers, rows = self.parse_quote_data(page, self.data_types)
        context.headers = headers
        context.rows = rows

    # Получаем ссылку на динамику курса
    def parse_currency_link(self, soup, currency):
        link = soup.find(text=currency).parent.attrs['href']
        return link

    # Получаем данные динамики курса
    def parse_quote_data(self, soup, types):
        headers = [cell.get_text(strip=True) for cell in soup.select('th')]

        rows = []
        for row in soup.select('tr'):
            cells = row.find_all('td')
            if not cells:
                continue
            row = [t(cell.get_text(strip=True)) for cell, t in zip(cells, types)]
            rows.append(row)

        return headers, rows

    # Получаем страницу
    def get_page_soup(self, url):
        # TODO Добавить обработку ошибок
        headers_response = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
        response = requests.get(url, headers=headers_response)
        soup = BeautifulSoup(response.text, features="lxml")
        sleep(1)
        return soup
