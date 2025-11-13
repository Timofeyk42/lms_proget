import sys
from PyQt6.QtCore import Qt, QStringListModel, QUrl, QSize
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit,
    QCompleter, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QFileDialog, QLabel, QListWidget, QListWidgetItem, QHeaderView
)
import requests
import xlsxwriter
import xml.etree.ElementTree as ET
import os
from datetime import datetime
import sqlite3
import ast
import operator

# =============== ЕДИНЫЙ СТИЛЬ ДЛЯ ВСЕГО ПРИЛОЖЕНИЯ ===============
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
QListView, QCompleter QListView {
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
}
QListWidget {
    background-color: #ffffff;
    border: 1px solid #c8e6c9;
    border-radius: 8px;
    padding: 4px;
}
QListWidget::item {
    padding: 8px;
    border-bottom: 1px solid #e0e0e0;
}
QListWidget::item:selected {
    background-color: #c8e6c9;
    color: #1b5e20;
}
"""

# =============== СЛОВАРЬ ПОЛНЫХ НАЗВАНИЙ ВАЛЮТ (НА РУССКОМ) ===============
CURRENCY_FULL_NAMES = {
    "AUD": "Австралийский доллар",
    "AZN": "Азербайджанский манат",
    "DZD": "Алжирский динар",
    "GBP": "Фунт стерлингов",
    "AMD": "Армянский драм",
    "BHD": "Бахрейнский динар",
    "BYN": "Белорусский рубль",
    "BGN": "Болгарский лев",
    "BOB": "Боливиано (Боливия)",
    "BRL": "Бразильский реал",
    "HUF": "Венгерский форинт",
    "VND": "Вьетнамский донг",
    "HKD": "Гонконгский доллар",
    "GEL": "Грузинский лари",
    "DKK": "Датская крона",
    "AED": "Дирхам ОАЭ",
    "USD": "Доллар США",
    "EUR": "Евро",
    "EGP": "Египетский фунт",
    "INR": "Индийская рупия",
    "IDR": "Индонезийская рупия",
    "IRR": "Иранский риал",
    "KZT": "Казахстанский тенге",
    "CAD": "Канадский доллар",
    "QAR": "Катарский риал",
    "KGS": "Киргизский сом",
    "CNY": "Китайский юань",
    "CUP": "Кубинское песо",
    "MDL": "Молдавский лей",
    "MNT": "Монгольский тугрик",
    "NGN": "Нигерийская найра",
    "NZD": "Новозеландский доллар",
    "NOK": "Норвежская крона",
    "OMR": "Оманский риал",
    "PLN": "Польский злотый",
    "SAR": "Саудовский риял",
    "RON": "Румынский лей",
    "XDR": "СДР (специальные права заимствования)",
    "SGD": "Сингапурский доллар",
    "TJS": "Таджикский сомони",
    "THB": "Тайский бат",
    "BDT": "Бангладешская така",
    "TRY": "Турецкая лира",
    "TMT": "Туркменский манат",
    "UZS": "Узбекский сум",
    "UAH": "Украинская гривна",
    "CZK": "Чешская крона",
    "SEK": "Шведская крона",
    "CHF": "Швейцарский франк",
    "ETB": "Эфиопский быр",
    "RSD": "Сербский динар",
    "ZAR": "Южноафриканский рэнд",
    "KRW": "Южнокорейская вона",
    "JPY": "Японская иена",
    "MMK": "Мьянманский кьят",
    "RUB": "Российский рубль",
}

SAFE_OPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
}

def format_currency_code(code):
    """Возвращает строку вида 'USD (US Dollar)' или просто 'RUB', если название неизвестно."""
    name = CURRENCY_FULL_NAMES.get(code, "")
    if name:
        return f"{code} ({name})"
    return code

def safe_eval(expr):
    """Безопасное вычисление простых арифметических выражений."""
    try:
        node = ast.parse(expr, mode='eval')
    except SyntaxError:
        raise ValueError("Некорректное выражение")

    def _eval(node):
        if isinstance(node, ast.Constant):
            return node.value
        elif isinstance(node, ast.Num):
            return node.n
        elif isinstance(node, ast.Expression):
            return _eval(node.body)
        elif isinstance(node, ast.UnaryOp):
            op = SAFE_OPS.get(type(node.op))
            if op is None:
                raise ValueError("Неподдерживаемая операция")
            return op(_eval(node.operand))
        elif isinstance(node, ast.BinOp):
            left = _eval(node.left)
            right = _eval(node.right)
            op = SAFE_OPS.get(type(node.op))
            if op is None:
                raise ValueError("Неподдерживаемая операция")
            return op(left, right)
        else:
            raise ValueError("Недопустимое выражение")

    return _eval(node.body)


def init_database():
    """Инициализация базы данных с проверкой целостности."""
    try:
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
    except sqlite3.Error as e:
        raise SystemExit("Не удалось инициализировать базу данных.")


def internet_connected(url='http://www.google.com', timeout=5):
    try:
        requests.head(url, timeout=timeout)
        return True
    except Exception:
        return False


def get_currency_rates_nn():
    currency_rates = {}
    conn = sqlite3.connect('curs_database.db')
    cursor = conn.cursor()
    data = cursor.execute("SELECT date FROM curss").fetchall()
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


# =============== Кастомный виджет для истории ===============
class wd1(QWidget):
    def __init__(self, arg2, arg3):
        super().__init__()
        self.setStyleSheet("background: transparent; border: none;")
        layout = QHBoxLayout()
        date_label = QLabel(arg2)
        date_label.setStyleSheet("font-weight: bold; color: #388e3c;")
        rate_label = QLabel(arg3)
        rate_label.setStyleSheet("color: #1b5e20;")
        layout.addWidget(date_label)
        layout.addStretch()
        layout.addWidget(rate_label)
        layout.setContentsMargins(10, 5, 10, 5)
        self.setLayout(layout)


# =============== Окно истории курсов ===============
class Wn3(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet(COMMON_STYLE)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("История курсов")
        self.setGeometry(1000, 250, 500, 500)

        self.lnval = QLineEdit()
        self.lnval.setPlaceholderText("Введите код валюты (например, USD)")

        self.btn = QPushButton("Показать историю")
        self.lstw = QListWidget()

        # Автодополнение
        self.curs = list(valcurss.keys())
        self.ln = QStringListModel(self.curs, self)
        self.compl = QCompleter()
        self.compl.setModel(self.ln)
        self.compl.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.lnval.setCompleter(self.compl)

        hb = QHBoxLayout()
        hb.addWidget(self.lnval)
        hb.addWidget(self.btn)

        vb = QVBoxLayout()
        vb.addLayout(hb)
        vb.addWidget(self.lstw)
        vb.setContentsMargins(15, 15, 15, 15)
        self.setLayout(vb)

        self.btn.clicked.connect(self.btnk)

    def btnk(self):
        self.lstw.clear()
        val = self.lnval.text().strip().upper()
        if not val:
            return
        conn = sqlite3.connect('curs_database.db')
        cursor = conn.cursor()
        rows = cursor.execute("SELECT curs, date FROM curss WHERE title = ? ORDER BY date DESC", (val,)).fetchall()
        conn.close()
        if not rows:
            item = QListWidgetItem("Нет данных")
            self.lstw.addItem(item)
            return
        seen = set()
        for curs, date in rows:
            if (curs, date) in seen:
                continue
            seen.add((curs, date))
            item = QListWidgetItem()
            item.setSizeHint(QSize(0, 60))
            self.lstw.addItem(item)
            formatted_code = format_currency_code(val)
            widget = wd1(date, f"{formatted_code}: {curs:.6f} RUB")
            self.lstw.setItemWidget(item, widget)


# =============== Окно таблицы всех курсов ===============
class Wn2(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet(COMMON_STYLE)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Все курсы")
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
        keys = sorted(valcurss.keys())
        self.tbl.setRowCount(len(keys))
        for i, key in enumerate(keys):
            self.tbl.setItem(i, 0, QTableWidgetItem(format_currency_code(key)))
            self.tbl.setItem(i, 1, QTableWidgetItem(f"{valcurss[key]:.6f}"))

    def xlsxxx(self, arg="Curss", path="/Users/denis/Downloads/"):
        os.makedirs(path, exist_ok=True)
        workbook = xlsxwriter.Workbook(f"{path}{arg}.xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Валюта")
        worksheet.write(0, 1, "Курс (RUB)")
        keys = sorted(valcurss.keys())
        for row, key in enumerate(keys, start=1):
            worksheet.write(row, 0, key)
            worksheet.write(row, 1, valcurss[key])
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


# =============== Основное окно конвертера ===============
class kalc(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet(COMMON_STYLE)
        self.v1 = self.v2 = self.f1 = ""
        try:
            with open("vals.txt", "r", encoding="utf-8") as f:
                self.v1 = f.readline().rstrip()
                self.v2 = f.readline().rstrip()
                self.f1 = f.readline().rstrip()
        except FileNotFoundError:
            pass
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Конвертер валют")
        self.setGeometry(250, 350, 650, 280)

        self.curs = list(valcurss.keys())
        self.curs_display = [format_currency_code(code) for code in self.curs]

        self.ln = QStringListModel(self.curs_display, self)
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

        left = QVBoxLayout()
        left.addWidget(QLabel("Количество:"))
        left.addWidget(self.frln)
        left.addWidget(QLabel("Результат:"))
        left.addWidget(self.scln)

        right = QVBoxLayout()
        right.addWidget(QLabel("Из валюты:"))
        right.addWidget(self.val1)
        right.addWidget(self.btswp)
        right.addWidget(QLabel("В валюту:"))
        right.addWidget(self.val2)
        right.addWidget(self.btn)

        main = QHBoxLayout()
        main.addLayout(left)
        main.addSpacing(30)
        main.addLayout(right)

        layout = QVBoxLayout()
        layout.addLayout(main)
        layout.setContentsMargins(25, 25, 25, 25)

        if not bl:
            lbl = QLabel(f"Последнее обновление: {bstd}")
            layout.addWidget(lbl)

        self.setLayout(layout)

    def closeEvent(self, event):
        try:
            v1_code = self.val1.text().strip().split()[0] if self.val1.text().strip() else ""
            v2_code = self.val2.text().strip().split()[0] if self.val2.text().strip() else ""
            with open("vals.txt", "w", encoding="utf-8") as f:
                f.write(f"{v1_code}\n")
                f.write(f"{v2_code}\n")
                f.write(f"{self.frln.text()}")
        except Exception:
            pass

        QApplication.quit()
        event.accept()

    def getvl(self):
        v1_text = self.val1.text().strip()
        v2_text = self.val2.text().strip()

        v1 = v1_text.split()[0] if v1_text else ""
        v2 = v2_text.split()[0] if v2_text else ""
        if not v1 or not v2:
            self.scln.setText("Укажите обе валюты")
            return
        if v1 not in valcurss or v2 not in valcurss:
            self.scln.setText("Неизвестная валюта")
            return
        try:
            try:
                val = safe_eval(self.frln.text().replace(' ', ''))
            except ValueError as e:
                self.scln.setText(f"Ошибка в выражении: {e}")
                return
            res = val * valcurss[v1] / valcurss[v2]
            self.scln.setText(f"{round(res, 4)}")
        except Exception as e:
            self.scln.setText(f"Ошибка: {e}")

    def swp(self):
        v1, v2 = self.val1.text().strip(), self.val2.text().strip()
        self.val1.setText(v2)
        self.val2.setText(v1)


# =============== Запуск приложения ===============
if __name__ == "__main__":
    app = QApplication(sys.argv)

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
    wn1 = Wn3()
    ex.show()
    wn.show()
    wn1.show()

    sys.exit(app.exec())