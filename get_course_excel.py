# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
from time import sleep


# Получаем ссылку на динамику курса
def parse_currency_link(soup, currency):
    link = soup.find(text=currency).parent.attrs['href']
    return link


# Получаем данные динамики курса
def parse_quote_data(soup):
    headers = [cell.get_text(strip=True) for cell in soup.select('th')]

    rows = []
    for row in soup.select('tr'):
        cells = row.find_all('td')
        if not cells:
            continue
        row = [cell.get_text(strip=True) for cell in cells]
        rows.append(row)

    return headers, rows


# Получаем страницу
def get_page_soup(url):
    # TODO Добавить обработку ошибок
    headers_response = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'}
    response = requests.get(url, headers=headers_response)
    soup = BeautifulSoup(response.text, features="lxml")
    sleep(1)
    return soup
