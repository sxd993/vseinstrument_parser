"""
Модуль для парсинга HTML-контента и извлечения данных о товарах.
Этот модуль:
- Загружает веб-страницы с помощью Selenium и парсит HTML с BeautifulSoup.
- Извлекает данные о товарах: артикул (только цифры), название, URL, цену, рейтинг и отзывы.
- Логирует ошибки и обновляет прогресс для интеграции с GUI.
"""

import asyncio
import re
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from utils.logger import logger

async def get_page_content(driver, url):
    """
    Загрузка страницы и возврат ее спарсенного HTML-контента.

    Аргументы:
        driver: Объект WebDriver Selenium.
        url (str): URL для загрузки.

    Возвращает:
        BeautifulSoup: Спарсенный HTML-контент или None, если загрузка не удалась.
    """
    try:
        logger.info(f"Загрузка страницы: {url}")
        driver.get(url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-qa="products-tile"]')))
        html = driver.page_source
        logger.info(f"Страница успешно загружена: {url}")
        return BeautifulSoup(html, "html.parser")
    except Exception as e:
        logger.error(f"Ошибка загрузки страницы {url}: {str(e)}")
        return None

async def parse_products(soup, max_products, current_product_count, progress_handler=None):
    """
    Извлечение данных о товарах из спарсенного HTML.

    Аргументы:
        soup: Объект BeautifulSoup с HTML страницы.
        max_products (int): Максимальное количество товаров для парсинга.
        current_product_count (int): Текущее количество спарсенных товаров.
        progress_handler: Объект для обновления прогресс-бара.

    Возвращает:
        tuple: Список данных о товарах и обновленное количество товаров.
    """
    products_data = []
    if not soup:
        return products_data, current_product_count

    products = soup.find_all("div", attrs={"data-qa": "products-tile"})
    if not products:
        logger.warning("Товары не найдены на странице")
        return products_data, current_product_count

    logger.info(f"Найдено товаров на странице: {len(products)}")
    for product in products:
        if max_products > 0 and current_product_count >= max_products:
            break
        try:
            # Извлечение артикула (только цифры)
            code_elem = product.find("p", attrs={"data-qa": "product-code-text"})
            code = "Н/Д"
            if code_elem:
                code_text = code_elem.text.strip()
                match = re.search(r'\d+', code_text)
                if match:
                    code = match.group()

            # Извлечение названия и URL
            name_elem = product.find("a", attrs={"data-qa": "product-name"})
            name = name_elem.text.strip() if name_elem else "Н/Д"
            url = (
                urljoin("https://www.vseinstrumenti.ru", name_elem["href"])
                if name_elem and name_elem.has_attr("href")
                else "Н/Д"
            )

            # Извлечение цены
            price_elem = product.find("p", attrs={"data-qa": "product-price-current"})
            price = price_elem.text.strip() if price_elem else "Н/Д"

            # Извлечение рейтинга и отзывов
            rating_container = product.find("a", attrs={"data-qa": "product-rating"})
            rating = "Н/Д"
            reviews = "Н/Д"

            if rating_container:
                rating_input = rating_container.find("input", attrs={"name": "rating"})
                if rating_input and rating_input.has_attr("value"):
                    try:
                        rating = f"{float(rating_input['value']):.2f}"
                    except ValueError:
                        pass

                reviews_elem = rating_container.find("span")
                if reviews_elem and reviews_elem.text.strip().isdigit():
                    reviews = reviews_elem.text.strip()

            products_data.append({
                "Артикул": code,
                "Название": name,
                "URL": url,
                "Цена": price,
                "Рейтинг": rating,
                "Отзывы": reviews,
            })

            current_product_count += 1
            if progress_handler:
                progress_handler.update(1)

        except Exception as e:
            logger.error(f"Ошибка парсинга товара: {str(e)}")
            continue

    return products_data, current_product_count