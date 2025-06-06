from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from urllib.parse import urljoin

# Metadata
current_date_time = "2025-06-06 08:04:00"
current_user_login = "sxd993"

def setup_driver():
    """Настройка Selenium-драйвера."""
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    )
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def get_page_content(driver, url):
    """Загрузка страницы и возврат HTML."""
    try:
        driver.get(url)
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-qa="products-tile"]'))
        )
        return BeautifulSoup(driver.page_source, 'html.parser')
    except:
        return None

def parse_products(soup, max_products, current_product_count):
    """Извлечение данных о товарах из HTML с учетом лимита."""
    products_data = []
    products = soup.find_all('div', attrs={'data-qa': 'products-tile'})
    
    if not products:
        return products_data, current_product_count

    for product in products:
        if max_products > 0 and current_product_count >= max_products:
            break
        try:
            # Извлечение артикула
            code_elem = product.find('p', attrs={'data-qa': 'product-code-text'})
            code = code_elem.text.strip() if code_elem else 'Артикул не найден'
            
            # Извлечение имени и URL товара
            name_elem = product.find('a', attrs={'data-qa': 'product-name'})
            name = name_elem.text.strip() if name_elem else 'Название не найдено'
            url = urljoin('https://www.vseinstrumenti.ru', name_elem['href']) if name_elem and name_elem.has_attr('href') else 'URL не найден'

            # Извлечение цены
            price_elem = product.find('p', attrs={'data-qa': 'product-price-current'})
            price = price_elem.text.strip() if price_elem else 'Цена не найдена'
            
            # Извлечение рейтинга
            rating_container = product.find('a', attrs={'data-qa': 'product-rating'})
            rating = 'Рейтинг не найден'
            reviews = 'Отзывы не найдены'
            
            if rating_container:
                rating_input = rating_container.find('input', attrs={'name': 'rating'})
                if rating_input and rating_input.has_attr('value'):
                    try:
                        rating = f"{float(rating_input['value']):.2f}"
                    except ValueError:
                        pass
                
                reviews_elem = rating_container.find('span')
                if reviews_elem and reviews_elem.text.strip().isdigit():
                    reviews = reviews_elem.text.strip()

            # Добавляем данные о товаре
            products_data.append({
                "Артикул": code,
                "Название": name,
                "URL": url,
                "Обычная цена": price,
                "Рейтинг": rating,
                "Количество отзывов": reviews,
            })

            current_product_count += 1

        except:
            continue
    
    return products_data, current_product_count

def get_next_page(driver, base_url, current_page=1):
    """Поиск и переход на следующую страницу."""
    try:
        next_button = driver.find_element(By.CSS_SELECTOR, 'a.next-page')
        if next_button.is_enabled() and next_button.is_displayed():
            next_button.click()
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-qa="products-tile"]'))
            )
            return driver.current_url
    except:
        next_page = current_page + 1
        next_url = f"{base_url.rstrip('/')}/page{next_page}/"
        try:
            driver.get(next_url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-qa="products-tile"]'))
            )
            return next_url
        except:
            return None

def save_to_excel(products_data, filename='products.xlsx'):
    """Сохранение данных в Excel."""
    try:
        df = pd.DataFrame(products_data)
        with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name="Товары", index=False)

        workbook = load_workbook(filename)
        worksheet = workbook["Товары"]

        for col_idx, col_name in enumerate(df.columns, start=1):
            column_letter = get_column_letter(col_idx)
            worksheet.column_dimensions[column_letter].width = 20

        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        workbook.save(filename)
    except Exception as e:
        raise

def main(url, max_products=0):
    """Основная функция парсинга."""
    driver = setup_driver()
    products_data = []
    current_product_count = 0
    base_url = url.split('#')[0]
    current_page = 1

    try:
        current_url = base_url
        while current_url and (max_products == 0 or current_product_count < max_products):
            soup = get_page_content(driver, current_url)
            if not soup:
                break

            page_products, current_product_count = parse_products(soup, max_products, current_product_count)
            products_data.extend(page_products)

            if max_products > 0 and current_product_count >= max_products:
                break

            current_url = get_next_page(driver, base_url, current_page)
            current_page += 1
            time.sleep(1)
    finally:
        driver.quit()

    if products_data:
        save_to_excel(products_data)
        print('Эксель создан')

if __name__ == "__main__":
    url = input("Введите URL для парсинга: ")
    max_products = int(input("Введите максимальное количество товаров (0 для всех): "))
    main(url, max_products)