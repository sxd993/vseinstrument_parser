"""
Модуль для настройки логирования.
Этот модуль:
- Инициализирует логгер с заданным форматом и уровнем логирования.
- Предоставляет единый логгер для использования во всех модулях приложения.
"""

import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)