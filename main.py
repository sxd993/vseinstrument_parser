"""
Основной скрипт для парсинга данных о товарах с сайта vseinstrumenti.ru.
Этот скрипт:
- Настраивает браузер Selenium для веб-скрейпинга.
- Рассчитывает количество страниц на основе 40 товаров на странице и запроса пользователя.
- Переходит по страницам с товарами, добавляя /pageX/ к URL.
- Извлекает данные о товарах с помощью BeautifulSoup.
- Сохраняет результаты в Excel-файл.
- Поддерживает отслеживание прогресса для интеграции с GUI.
Скрипт обрабатывает ошибки и логирует ключевые события.
"""

import asyncio
from urllib.parse import urlparse, parse_qs, urlencode
from utils.selenium_driver import setup_browser
from utils.excel_creator import save_to_excel
from utils.parse import get_page_content, parse_products
from utils.logger import logger

async def main(url, max_products=0, progress_handler=None, output_file="products.xlsx"):
    """
    Основная функция парсинга.

    Аргументы:
        url (str): URL страницы категории товаров для парсинга.
        max_products (int): Максимальное количество товаров для парсинга (0 для всех).
        progress_handler: Объект для обновления прогресс-бара в GUI.
        output_file (str): Имя выходного Excel-файла.

    Возвращает:
        None
    """
    driver = setup_browser()
    products_data = []
    current_product_count = 0
    base_url = url.split("#")[0]
    current_page = 1

    # Рассчитываем количество страниц (40 товаров на странице)
    total_pages = (max_products + 39) // 40 if max_products > 0 else float('inf')
    logger.info(f"Запрошено {max_products} товаров, требуется {total_pages} страниц")

    if progress_handler and max_products > 0:
        progress_handler.set_total(max_products)

    try:
        current_url = base_url
        while current_url and (max_products == 0 or current_product_count < max_products):
            if current_page > total_pages:
                logger.info(f"Достигнуто максимальное количество страниц: {total_pages}")
                break

            soup = await get_page_content(driver, current_url)
            if not soup:
                logger.warning(f"Не удалось загрузить страницу: {current_url}")
                break

            page_products, current_product_count = await parse_products(
                soup, max_products, current_product_count, progress_handler
            )
            products_data.extend(page_products)

            if max_products > 0 and current_product_count >= max_products:
                logger.info(f"Достигнуто запрошенное количество товаров: {max_products}")
                break

            # Формируем URL следующей страницы
            parsed_url = urlparse(base_url)
            query = parse_qs(parsed_url.query)
            next_page_num = current_page + 1
            next_path = parsed_url.path.rstrip('/') + f'/page{next_page_num}/'
            next_url = parsed_url._replace(path=next_path).geturl()
            if query:
                next_url += '?' + urlencode(query, doseq=True)

            current_url = next_url if next_page_num <= total_pages else None
            current_page += 1
            await asyncio.sleep(0.5)  # Задержка для избежания блокировки

    except Exception as e:
        logger.error(f"Ошибка парсинга: {str(e)}")
    finally:
        driver.quit()

    if products_data:
        save_to_excel(products_data, output_file)
        logger.info(f"Парсинг завершен, Excel сохранен: {output_file}")
    else:
        logger.warning("Нет данных для сохранения")

if __name__ == "__main__":
    url = input("Введите URL для парсинга (например, https://www.vseinstrumenti.ru/category/perforatory-32/): ")
    max_products = int(input("Введите максимальное количество товаров (0 для всех): "))
    asyncio.run(main(url, max_products))