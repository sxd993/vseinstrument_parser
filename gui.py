"""
Модуль графического интерфейса для парсера vseinstrumenti.ru с использованием PyQt5.
Этот модуль:
- Предоставляет удобный интерфейс для ввода URL, максимального количества товаров и имени файла.
- Отображает прогресс парсинга через прогресс-бар.
- Показывает статус парсинга и ошибки в текстовом поле.
- Интегрируется с логикой парсинга через asyncio и qasync.
- Использует темную тему для лучшей читаемости.
"""

import sys
import asyncio
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QProgressBar
)
from PyQt5.QtCore import Qt, pyqtSignal, QObject
from PyQt5 import QtGui
import qasync
from main import main
from utils.logger import logger

class TqdmToProgressBar(QObject):
    """
    Класс для перенаправления обновлений прогресса на QProgressBar.
    """
    progress_updated = pyqtSignal(int)
    total_updated = pyqtSignal(int)

    def __init__(self, progress_bar):
        super().__init__()
        self.progress_bar = progress_bar
        self.progress_updated.connect(self.progress_bar.setValue)
        self.total_updated.connect(self.progress_bar.setMaximum)

    def update(self, n=1):
        self.progress_updated.emit(self.progress_bar.value() + n)

    def set_total(self, total):
        self.total_updated.emit(total)

class ParserApp(QMainWindow):
    """
    Главное окно приложения парсера.
    """
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        """
        Инициализация пользовательского интерфейса.
        """
        self.setWindowTitle("Парсер Vseinstrumenti.ru")
        self.setGeometry(100, 100, 450, 520)
        self.setStyleSheet("background-color: #333333;")

        # Основной виджет и компоновка
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setAlignment(Qt.AlignCenter)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Заголовок
        title_label = QLabel("Парсер Vseinstrumenti.ru")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #FFFFFF;")
        title_layout = QHBoxLayout()
        title_layout.addStretch()
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        main_layout.addLayout(title_layout)

        # Поле ввода URL
        self.url_label = QLabel("URL для парсинга:")
        self.url_label.setStyleSheet("font-size: 14px; color: #DDDDDD;")
        self.url_example = QLabel("Например, https://www.vseinstrumenti.ru/category/perforatory-32/")
        self.url_example.setStyleSheet("font-size: 12px; color: #BBBBBB; font-style: italic;")
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("Введите URL")
        self.url_input.setStyleSheet(
            "font-size: 14px; padding: 5px; border: 1px solid #555555; "
            "border-radius: 5px; background-color: #444444; color: #FFFFFF;"
        )
        url_layout = QHBoxLayout()
        url_layout.addStretch()
        url_layout.addWidget(self.url_input, 1)
        url_layout.addStretch()
        main_layout.addWidget(self.url_label)
        main_layout.addWidget(self.url_example)
        main_layout.addLayout(url_layout)

        # Поле ввода максимального количества товаров
        self.max_products_label = QLabel("Максимальное количество товаров (0 для всех):")
        self.max_products_label.setStyleSheet("font-size: 14px; color: #DDDDDD;")
        self.max_products_input = QLineEdit("0")
        self.max_products_input.setValidator(QtGui.QIntValidator(0, 1000000))
        self.max_products_input.setStyleSheet(
            "font-size: 14px; padding: 5px; border: 1px solid #555555; "
            "border-radius: 5px; background-color: #444444; color: #FFFFFF;"
        )
        max_products_layout = QHBoxLayout()
        max_products_layout.addStretch()
        max_products_layout.addWidget(self.max_products_input, 1)
        max_products_layout.addStretch()
        main_layout.addWidget(self.max_products_label)
        main_layout.addLayout(max_products_layout)

        # Поле ввода имени выходного файла
        self.output_file_label = QLabel("Имя выходного файла:")
        self.output_file_label.setStyleSheet("font-size: 14px; color: #DDDDDD;")
        self.output_file_input = QLineEdit("products.xlsx")
        self.output_file_input.setStyleSheet(
            "font-size: 14px; padding: 5px; border: 1px solid #555555; "
            "border-radius: 5px; background-color: #444444; color: #FFFFFF;"
        )
        output_file_layout = QHBoxLayout()
        output_file_layout.addStretch()
        output_file_layout.addWidget(self.output_file_input, 1)
        output_file_layout.addStretch()
        main_layout.addWidget(self.output_file_label)
        main_layout.addLayout(output_file_layout)

        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                font-size: 12px; padding: 5px; border: 1px solid #555555;
                border-radius: 5px; background-color: #444444; color: #FFFFFF;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
        """)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        progress_layout = QHBoxLayout()
        progress_layout.addStretch()
        progress_layout.addWidget(self.progress_bar, 1)
        progress_layout.addStretch()
        main_layout.addLayout(progress_layout)

        # Кнопка запуска парсинга
        self.parse_button = QPushButton("Начать парсинг")
        self.parse_button.setStyleSheet("""
            QPushButton {
                font-size: 16px; padding: 10px; background-color: #4A90E2;
                color: white; border: none; border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #357ABD;
            }
            QPushButton:disabled {
                background-color: #A0C4FF;
            }
        """)
        parse_button_layout = QHBoxLayout()
        parse_button_layout.addStretch()
        parse_button_layout.addWidget(self.parse_button, 1)
        parse_button_layout.addStretch()
        self.parse_button.clicked.connect(self.start_parsing)
        main_layout.addLayout(parse_button_layout)

        # Поле вывода статуса
        self.status_output = QTextEdit()
        self.status_output.setReadOnly(True)
        self.status_output.setStyleSheet(
            "font-size: 12px; border: 1px solid #555555; border-radius: 5px; "
            "padding: 5px; background-color: #444444; color: #FFFFFF;"
        )
        self.status_output.setFixedHeight(100)
        status_layout = QHBoxLayout()
        status_layout.addStretch()
        status_layout.addWidget(self.status_output, 1)
        status_layout.addStretch()
        main_layout.addLayout(status_layout)

    async def run_parsing(self, url, max_products, output_file, progress_handler):
        """
        Запуск процесса парсинга.

        Аргументы:
            url (str): URL для парсинга.
            max_products (int): Максимальное количество товаров.
            output_file (str): Имя выходного Excel-файла.
            progress_handler: Объект для обновления прогресс-бара.
        """
        try:
            await main(url, max_products, progress_handler, output_file)
            self.status_output.append(f"Парсинг завершен. Файл сохранен: {output_file}")
        except Exception as e:
            self.status_output.append(f"Ошибка парсинга: {str(e)}")
        finally:
            self.parse_button.setEnabled(True)
            self.progress_bar.setValue(0)

    @qasync.asyncSlot()
    async def start_parsing(self):
        """
        Запуск парсинга при нажатии кнопки.
        """
        url = self.url_input.text().strip()
        if not url:
            self.status_output.append("Ошибка: Введите URL для парсинга")
            return
        try:
            max_products = int(self.max_products_input.text())
        except ValueError:
            self.status_output.append("Ошибка: Введите корректное число товаров")
            return
        output_file = self.output_file_input.text().strip()
        if not output_file:
            self.status_output.append("Ошибка: Введите имя выходного файла")
            return
        self.parse_button.setEnabled(False)
        self.status_output.append("Парсинг начат...")
        progress_handler = TqdmToProgressBar(self.progress_bar)
        await self.run_parsing(url, max_products, output_file, progress_handler)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)
    window = ParserApp()
    window.show()
    with loop:
        loop.run_forever()