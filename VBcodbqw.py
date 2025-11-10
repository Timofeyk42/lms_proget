import sys
from PyQt6.QtCore import Qt, QStringListModel, QUrl
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit,
    QCompleter, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QFileDialog, QLabel, QHeaderView
)
import requests
import xlsxwriter
import xml.etree.ElementTree as ET
import os
from datetime import datetime
import sqlite3


# Единый мягкий зелёный стиль
COMMON_STYLE = """
QWidget {
    background-color: #f0f9f1;
    color: #1b5e20;
    font-family: 'Segoe UI', sans-serif;
    font-size: 14px;
}
QLineEdit {
    background-color: #e8f5eb;
    border: 1px solid #81c784;
    border-radius: 8px;
    padding: 10px;
    color: #1b5e20;
    selection-background-color: #a5d6a7;
}
QLineEdit:read-only {
    background-color: #c8e6c9;
    font-weight: bold;
    border-style: dashed;
}
QPushButton {
    background-color: #66bb6a;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px;
    font-weight: bold;
    font-size: 14px;
}
QPushButton:hover {
    background-color: #4caf50;
}
QListView {
    background-color: #ffffff;
    color: #1b5e20;
    border: 1px solid #81c784;
    selection-background-color: #a5d6a7;
}
QTableWidget {
    background-color: #ffffff;
    gridline-color: #c8e6c9;
    alternate-background-color: #f1f8e9;
}
QHeaderView::section {
    background-color: #e8f5eb;
    color: #1b5e20;
    padding: 6px;
    border: 1px solid #c8e6c9;
    font-weight: bold;
}
QLabel {
    color: #388e3c;
    font-size: 13px;
}
"""


def internet_connected(url='http://www.google.com', timeout=5):
    try:
        requests.head(url, timeout=timeout)
        return True
    except requests.ConnectionError:
        return False


def get_currency_rates_nn():
    currency_rates = {}
    conn = sqlite3.connect('curs_database.db')
    cursor = conn.cursor()
    data = cursor.execute("SELECT date FROM curss").fetchall()
    if not data:
        conn.close()
        return {}, ""
    bstdt = sorted(data, key=lambda x: x[0], reverse=True)[0][0]
    curss = cursor.execute("SELECT title, curs FROM curss WHERE date = ?", (bstdt,)).fetchall()
    conn.close()
    for i in curss:
        currency_rates[i[0]] = i[1]
    return currency_rates, bstdt


def get_currency_rates():
    url = "https://www.cbr.ru/scripts/XML_daily.asp"
    response = requests.get(url)
    response.raise_for_status()

    root = ET.fromstring(response.content)
    currency_rates = {}

    conn = sqlite3.connect('curs_database.db')
    cursor = conn.cursor()
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    for valute in root.findall('Valute'):
        char_code = valute.find('CharCode').text
        nominal = int(valute.find('Nominal').text)
        value = float(valute.find('Value').text.replace(',', '.'))
        rate = value / nominal
        currency_rates[char_code] = rate
        cursor.execute("INSERT INTO curss (title, curs, date) VALUES (?, ?, ?)",
                       (char_code, rate, current_time))

    currency_rates["RUB"] = 1.0
    conn.commit()
    conn.close()
    return currency_rates


class Wn2(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet(COMMON_STYLE)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Курс валют")
        self.setGeometry(850, 250, 550, 500)

        self.uplbtn = QPushButton("Скачать как .xlsx")
        self.tbl = QTableWidget()

        layout = QVBoxLayout()
        layout.addWidget(self.tbl)
        layout.addWidget(self.uplbtn)
        layout.setContentsMargins(20, 20, 20, 20)
        self.setLayout(layout)

        self.tbl.setColumnCount(2)
        self.tbl.setHorizontalHeaderLabels(["Валюта", "Курс (RUB)"])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl.setAlternatingRowColors(True)
        self.tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        self.tblz()
        self.uplbtn.clicked.connect(self.show_save_file_dialog)

    def tblz(self):
        self.tbl.setRowCount(len(valcurss))
        c = 0
        for key in sorted(valcurss.keys()):
            self.tbl.setItem(c, 0, QTableWidgetItem(key))
            self.tbl.setItem(c, 1, QTableWidgetItem(f"{valcurss[key]:.6f}"))
            c += 1

    def xlsxxx(self, arg="Curss", path="/Users/denis/Downloads/"):
        a = valcurss
        os.makedirs(path, exist_ok=True)
        workbook = xlsxwriter.Workbook(f"{path}{arg}.xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Валюта")
        worksheet.write(0, 1, "Курс (RUB)")
        row = 1
        for key in sorted(a):
            worksheet.write(row, 0, key)
            worksheet.write(row, 1, a[key])
            row += 1
        workbook.close()

    def show_save_file_dialog(self):
        url, _ = QFileDialog.getSaveFileUrl(
            parent=self,
            caption="Сохранить файл",
            filter="Excel Files (*.xlsx)",
            initialFilter="Excel Files (*.xlsx)"
        )
        if url.isEmpty():
            return

        file_path = url.toLocalFile()
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

        dir_path = os.path.dirname(file_path) + os.sep
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        self.xlsxxx(arg=base_name, path=dir_path)


class kalc(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet(COMMON_STYLE)
        # Загрузка сохранённых значений
        self.v1 = self.v2 = self.f1 = ""
        try:
            with open("vals.txt", "r", encoding="utf-8") as f:
                self.v1 = f.readline().rstrip()
                self.v2 = f.readline().rstrip()
                self.f1 = f.readline().rstrip()
        except FileNotFoundError:
            pass  # Файл не существует — используем пустые значения
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Конвертер валют")
        self.setGeometry(250, 350, 650, 280)

        self.curs = list(valcurss.keys())
        self.ln = QStringListModel(self.curs, self)
        self.compl = QCompleter()
        self.compl.setModel(self.ln)
        self.compl.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

        self.frln = QLineEdit(self.f1)
        self.frln.setPlaceholderText("Сумма (например: 100 или 50*1.2)")

        self.scln = QLineEdit()
        self.scln.setPlaceholderText("Результат")
        self.scln.setReadOnly(True)

        self.val1 = QLineEdit(self.v1)
        self.val1.setPlaceholderText("Из валюты (например, USD)")
        self.val2 = QLineEdit(self.v2)
        self.val2.setPlaceholderText("В валюту (например, EUR)")

        self.val1.setCompleter(self.compl)
        self.val2.setCompleter(self.compl)

        self.btn = QPushButton("Конвертировать")
        self.btn.clicked.connect(self.getvl)

        self.btswp = QPushButton("⇄")
        self.btswp.setFixedWidth(50)
        self.btswp.clicked.connect(self.swp)

        # Левая колонка
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Сумма:"))
        left_layout.addWidget(self.frln)
        left_layout.addWidget(QLabel("Результат:"))
        left_layout.addWidget(self.scln)

        # Правая колонка
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("Из валюты:"))
        right_layout.addWidget(self.val1)
        right_layout.addWidget(self.btswp)
        right_layout.addWidget(QLabel("В валюту:"))
        right_layout.addWidget(self.val2)
        right_layout.addWidget(self.btn)

        main_layout = QHBoxLayout()
        main_layout.addLayout(left_layout)
        main_layout.addSpacing(30)
        main_layout.addLayout(right_layout)

        layout = QVBoxLayout()
        layout.addLayout(main_layout)
        layout.setContentsMargins(25, 25, 25, 25)

        if not bl:
            self.lbl = QLabel(f"Последнее обновление курсов: {bstd}")
            layout.addWidget(self.lbl)

        self.setLayout(layout)

    def closeEvent(self, event):
        try:
            with open("vals.txt", "w", encoding="utf-8") as f:
                f.write(f"{self.val1.text().upper()}\n")
                f.write(f"{self.val2.text().upper()}\n")
                f.write(f"{self.frln.text()}")
        except Exception:
            pass  # Игнорируем ошибки записи
        event.accept()

    def getvl(self):
        vl1 = self.val1.text().strip().upper()
        vl2 = self.val2.text().strip().upper()
        if not vl1 or not vl2:
            self.scln.setText("Укажите обе валюты")
            return
        if vl1 not in valcurss or vl2 not in valcurss:
            self.scln.setText("Неизвестная валюта")
            return
        try:
            val = eval(self.frln.text(), {"__builtins__": {}})
            result = val * valcurss[vl1] / valcurss[vl2]
            self.scln.setText(f"{result:.6f}")
        except Exception as e:
            self.scln.setText(f"Ошибка: {e}")

    def swp(self):
        v1 = self.val1.text().strip()
        v2 = self.val2.text().strip()
        self.val1.setText(v2)
        self.val2.setText(v1)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Инициализация БД
    conn = sqlite3.connect('curs_database.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS curss (
            title TEXT NOT NULL,
            curs REAL NOT NULL,
            date TEXT NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS names (
            name TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

    # Загрузка курсов
    bl = internet_connected()
    if bl:
        valcurss = get_currency_rates()
        bstd = datetime.now().strftime("%Y-%m-%d %H:%M")
    else:
        valcurss, bstd = get_currency_rates_nn()
        if not valcurss:
            valcurss = {"RUB": 1.0}
            bstd = "нет данных"

    ex = kalc()
    wn = Wn2()
    ex.show()
    wn.show()

    sys.exit(app.exec())