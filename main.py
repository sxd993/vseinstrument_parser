from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import logging
from urllib.parse import urljoin
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def setup_driver():
    """Настройка Selenium-драйвера."""
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Фоновый режим
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def get_page_content(driver, url):
    """Загрузка страницы и возврат HTML."""
    try:
        logging.info(f"Загрузка страницы: {url}")
        driver.get(url)
        time.sleep(3)  # Ждем загрузки динамического контента
        return BeautifulSoup(driver.page_source, 'html.parser')
    except Exception as e:
        logging.error(f"Ошибка при загрузке страницы {url}: {e}")
        return None

def parse_products(soup, max_products, current_product_count):
    """Извлечение данных о товарах из HTML с учетом лимита."""
    products_data = []
    supplier_data = set()  # Для уникальных продавцов
    products = soup.find_all('div', attrs={'data-qa': 'products-tile'})
    
    if not products:
        logging.warning("Товары не найдены на странице")
        return products_data, supplier_data, current_product_count

    for product in products:
        if max_products > 0 and current_product_count >= max_products:
            break
        try:
            # Извлечение данных
            code = product.find('p', attrs={'data-qa': 'product-code-text'}).text.strip() if product.find('p', attrs={'data-qa': 'product-code-text'}) else 'N/A'
            name = product.find('a', attrs={'data-qa': 'product-name'}).text.strip() if product.find('a', attrs={'data-qa': 'product-name'}) else 'N/A'
            url = product.find('a', attrs={'data-qa': 'product-name'})['href'] if product.find('a', attrs={'data-qa': 'product-name'}) else 'N/A'
            price = product.find('p', attrs={'data-qa': 'product-price-current'}).text.strip() if product.find('p', attrs={'data-qa': 'product-price-current'}) else 'N/A'
            old_price = product.find('span', attrs={'data-qa': 'product-price-old-value'}).text.strip() if product.find('span', attrs={'data-qa': 'product-price-old-value'}) else 'N/A'
            discount = product.find('span', attrs={'data-qa': 'product-price-discount'}).text.strip() if product.find('span', attrs={'data-qa': 'product-price-discount'}) else 'N/A'
            availability = product.find('p', attrs={'data-qa': 'product-availability-total-available'}).text.strip() if product.find('p', attrs={'data-qa': 'product-availability-total-available'}) else 'N/A'
            rating = product.find('input', attrs={'name': 'rating'})['value'] if product.find('input', attrs={'name': 'rating'}) else 'N/A'

            # Поля, которые сложно извлечь из текущего HTML (заглушки)
            brand = 'N/A'  # Бренд не указан в HTML, можно извлечь из названия
            supplier = 'N/A'  # Название продавца требует доп. парсинга
            supplier_id = 'N/A'  # ID продавца требует доп. парсинга
            supplier_rating = 'N/A'  # Рейтинг продавца требует доп. парсинга

            # Добавляем данные о товаре
            products_data.append({
                "Артикул": code,
                "Название": name,
                "Бренд": brand,
                "URL": url,
                "Обычная цена": price,
                "Цена по ВБ Карте": 'N/A',  # Заглушка
                "Цена без ВБ Карте": 'N/A',  # Заглушка
                "Отзывы": 'N/A',  # Кол-во отзывов не найдено
                "Рейтинг": rating,
                "Поставщик(продавец)": supplier,
                "ID продавца": supplier_id,
                "Рейтинг продавца": supplier_rating,
            })

            # Добавляем данные о продавце (заглушка, если нет данных)
            supplier_data.add((
                supplier_id,
                supplier,
                'N/A',  # Полное юридическое название
                'N/A',  # ИНН
                'N/A',  # ОГРН
                'N/A',  # ОГРНИП
                'N/A',  # Юридический адрес
                'N/A',  # Торговая марка
                'N/A',  # КПП
                'N/A',  # Номер регистрации
                'N/A',  # УНП
                'N/A',  # БИН
                'N/A',  # УНН
                'N/A',  # Ссылка на продавца
            ))

            current_product_count += 1
            logging.info(f"Обработан товар: {name}, Код: {code}, Цена: {price}")

        except AttributeError as e:
            logging.error(f"Ошибка при извлечении данных для товара: {e}")
            continue
    
    return products_data, supplier_data, current_product_count

def get_next_page(soup, base_url):
    """Поиск URL следующей страницы."""
    next_page = soup.find('a', class_='next-page')
    if next_page and 'href' in next_page.attrs:
        next_url = next_page['href']
        if not next_url.startswith('http'):
            next_url = urljoin(base_url, next_url)
        logging.info(f"Найдена следующая страница: {next_url}")
        return next_url
    logging.info("Пагинация завершена")
    return None

def save_to_excel(products_data, supplier_data, filename='products.xlsx'):
    """Сохранение данных в Excel в стиле excel_creator.py."""
    try:
        # Преобразуем данные о товарах в DataFrame
        df = pd.DataFrame(products_data)

        # Преобразуем данные о продавцах в DataFrame
        supplier_df = pd.DataFrame(
            list(supplier_data),
            columns=[
                "ID продавца", "Название продавца", "Полное юридическое название",
                "ИНН", "ОГРН", "ОГРНИП", "Юридический адрес", "Торговая марка",
                "КПП", "Номер регистрации", "УНП", "БИН", "УНН", "Ссылка на продавца"
            ]
        )

        # Сохраняем в Excel с несколькими листами
        with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name="Товары", index=False)
            supplier_df.to_excel(writer, sheet_name="Продавцы", index=False)

        # Настройка ширины столбцов
        workbook = load_workbook(filename)
        default_width = 8.43
        first_width = float(default_width * 1.3)
        second_width = float(default_width * 3)
        third_width = float(default_width * 6)

        # Лист Товары
        worksheet = workbook["Товары"]
        columns_to_widen_one = ["Артикул", "Бренд", "Обычная цена", "Цена по ВБ Карте", "Цена без ВБ Карте", "Отзывы", "Рейтинг", "Рейтинг продавца"]
        columns_to_widen_two = ["Поставщик(продавец)", "ID продавца"]
        columns_to_widen_three = ["Название", "URL"]
        for col_idx, col_name in enumerate(df.columns, start=1):
            column_letter = get_column_letter(col_idx)
            if col_name in columns_to_widen_one:
                worksheet.column_dimensions[column_letter].width = first_width
                logging.debug(f"Установлена ширина {first_width} для столбца {col_name} ({column_letter})")
            elif col_name in columns_to_widen_two:
                worksheet.column_dimensions[column_letter].width = second_width
                logging.debug(f"Установлена ширина {second_width} для столбца {col_name} ({column_letter})")
            elif col_name in columns_to_widen_three:
                worksheet.column_dimensions[column_letter].width = third_width
                logging.debug(f"Установлена ширина {third_width} для столбца {col_name} ({column_letter})")

        # Лист Продавцы
        supplier_worksheet = workbook["Продавцы"]
        supplier_columns_to_widen_one = ["ID продавца", "ИНН", "ОГРН", "ОГРНИП", "КПП", "Номер регистрации", "УНП", "БИН", "УНН"]
        supplier_columns_to_widen_two = ["Название продавца", "Торговая марка"]
        supplier_columns_to_widen_three = ["Полное юридическое название", "Юридический адрес", "Ссылка на продавца"]
        for col_idx, col_name in enumerate(supplier_df.columns, start=1):
            column_letter = get_column_letter(col_idx)
            if col_name in supplier_columns_to_widen_one:
                supplier_worksheet.column_dimensions[column_letter].width = first_width
                logging.debug(f"Установлена ширина {first_width} для столбца {col_name} ({column_letter}) в листе Продавцы")
            elif col_name in supplier_columns_to_widen_two:
                supplier_worksheet.column_dimensions[column_letter].width = second_width
                logging.debug(f"Установлена ширина {second_width} для столбца {col_name} ({column_letter}) в листе Продавцы")
            elif col_name in supplier_columns_to_widen_three:
                supplier_worksheet.column_dimensions[column_letter].width = third_width
                logging.debug(f"Установлена ширина {third_width} для столбца {col_name} ({column_letter}) в листе Продавцы")

        # Форматирование заголовков
        for sheet in [worksheet, supplier_worksheet]:
            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')

        workbook.save(filename)
        logging.info(f"Excel файл успешно создан: {filename}")

    except Exception as e:
        logging.error(f"Ошибка при записи Excel файла {filename}: {str(e)}")
        raise

def main(url, max_products=0):
    """Основная функция парсинга с лимитом товаров."""
    driver = setup_driver()
    products_data = []
    supplier_data = set()
    current_product_count = 0
    base_url = url.split('#')[0]  # Убираем хэш-параметры

    try:
        current_url = base_url
        while current_url and (max_products == 0 or current_product_count < max_products):
            soup = get_page_content(driver, current_url)
            if not soup:
                break

            # Парсинг товаров
            page_products, page_suppliers, current_product_count = parse_products(soup, max_products, current_product_count)
            products_data.extend(page_products)
            supplier_data.update(page_suppliers)

            # Проверка лимита товаров
            if max_products > 0 and current_product_count >= max_products:
                logging.info(f"Достигнут лимит товаров: {max_products}")
                break

            # Поиск следующей страницы
            current_url = get_next_page(soup, base_url)
            time.sleep(3)  # Задержка для этичного парсинга

    finally:
        driver.quit()

    # Сохранение в Excel
    if products_data:
        save_to_excel(products_data, supplier_data)
    else:
        logging.warning("Нет данных для сохранения")

if __name__ == "__main__":
    url = input("Введите URL для парсинга (например, https://www.vseinstrumenti.ru/category/perforatory-32/): ")
    max_products = int(input("Введите максимальное количество товаров для парсинга (0 для всех): "))
    main(url, max_products)