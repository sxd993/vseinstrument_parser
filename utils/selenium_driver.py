"""
Модуль для настройки браузера Selenium.
Этот модуль:
- Инициализирует headless-браузер Chrome с реалистичным user-agent.
- Использует webdriver-manager для автоматической установки ChromeDriver.
- Возвращает объект WebDriver для использования в парсинге.
- Минимизирует логирование, выводя только ключевые сообщения.
"""

import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from utils.logger import logger

# Отключаем логирование webdriver-manager
os.environ["WDM_LOG"] = "0"


def setup_browser():
    """
    Настройка экземпляра браузера Selenium.

    Возвращает:
        WebDriver: Объект браузера Selenium.
    """
    try:
        logger.info("Запуск Selenium")
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        )
        options.add_argument("--log-level=3")  # Минимизируем логи ChromeDriver
        service = Service(ChromeDriverManager().install(), log_output=os.devnull)
        driver = webdriver.Chrome(service=service, options=options)
        logger.info("Selenium успешно запущен")
        return driver
    except Exception as e:
        logger.error(f"Ошибка настройки Selenium: {str(e)}")
        raise
