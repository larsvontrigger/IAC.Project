import os
import pandas as pd
from datetime import datetime
from PyQt5 import QtWidgets, QtCore
import locale
import unicodedata
from iac.database import details_paths, root_measurment, root_formular  # Импортируем базу данных

class ShiftApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # Работа с локалями
        locale.setlocale(locale.LC_TIME, 'sk_SK.UTF-8')
     
        # Настройка интерфейса
        self.setWindowTitle('Vyber si smenu')
        self.setGeometry(300, 300, 600, 400)

        # Выпадающий список для смен
        self.shift_label = QtWidgets.QLabel('Vyber si smenu:', self)
        self.shift_label.move(20, 26)

        self.shift_combobox = QtWidgets.QComboBox(self)
        self.shift_combobox.addItems(['Ranná', 'Poobedna', 'Nočna'])
        self.shift_combobox.move(120, 20)

        # Кнопка для запуска процесса
        self.process_button = QtWidgets.QPushButton('Ukazať diely', self)
        self.process_button.move(250, 20)
        self.process_button.clicked.connect(self.on_select_shift)

        # Поле для вывода результатов (QTextBrowser поддерживает ссылки)
        self.result_text = QtWidgets.QTextBrowser(self)
        self.result_text.setGeometry(20, 60, 560, 300)
        self.result_text.setOpenExternalLinks(True)  # Разрешаем открытие ссылок

    @staticmethod
    def remove_diacritics(input_str):
        # Преобразуем строку в нормализованный вид (NFD разбивает символы с диакритиками)
        nfkd_form = unicodedata.normalize('NFKD', input_str)
        # Убираем все символы, которые не являются символами ASCII (диакритики удаляются)
        return ''.join([char for char in nfkd_form if not unicodedata.combining(char)])

    @staticmethod
    def create_hyperlink(file_path, display_name):
        # Генерация HTML-ссылки для файла
        return f'<a href="file://{file_path}">{display_name}</a>'

    def on_select_shift(self):
        shift = self.shift_combobox.currentText()
        self.process_shift(shift)

    def process_shift(self, shift):
        today = datetime.today()
        current_day = today.strftime("%d.%B").lower()
        current_month = today.strftime("%m %B").lower()
        
        # Убираем 0 из current_day
        if current_day.startswith('0'):
            current_day = current_day[1:]

        # Удаляем диакритические знаки из текущего месяца
        current_month = self.remove_diacritics(current_month)
        current_day = self.remove_diacritics(current_day)
        
        folder_path_base = "/Users/piratejet/Documents/IAC.SERVER/Logistics/02 - PLANY VYROBY/Plany vyroby IM/2024"
        month_folder_path = os.path.join(folder_path_base, current_month)

        file_name = None
        if os.path.exists(month_folder_path):
            for file in os.listdir(month_folder_path):
                if current_day in file.lower() and file.endswith('.xlsx'):
                    file_name = file
                    break
        else:
            self.result_text.setText(f"Subor {current_month} neexistuje.")
            return

        if file_name:
            file_path = os.path.join(month_folder_path, file_name)
            data = pd.read_excel(file_path, sheet_name='INJECTION MOULDING')
        else:
            self.result_text.setText(f"Plan na {current_day} ešte neexistuje.")
            return

        parts_for_shift_indices = self.get_parts_for_shift(data, shift)
        project_parts = self.get_project_with_parts(data, parts_for_shift_indices)

        self.result_text.clear()
        for project, parts in project_parts.items():
            self.result_text.append(f"\n{project}:")

            for part in parts:
                # Проверяем, есть ли деталь в базе данных
                if part in details_paths:
                    # Получаем путь к таблице с измерениями
                    measurment_table = details_paths[part].get('measurment_table')
                    if measurment_table:
                        measurment_path = os.path.join(root_measurment, measurment_table[0])
                        # Создаем гиперссылку
                        hyperlink = self.create_hyperlink(measurment_path, part)
                        self.result_text.append(f"  Измерения: {hyperlink}")

                    # Получаем путь к формуляру
                    formular_table = details_paths[part].get('formular_table')
                    if formular_table:
                        formular_path = os.path.join(root_formular, formular_table[0])
                        # Создаем гиперссылку
                        hyperlink = self.create_hyperlink(formular_path, part)
                        self.result_text.append(f"  Формуляр: {hyperlink}")
                else:
                    self.result_text.append(f"  {part}")

    def get_parts_for_shift(self, data, shift):
        if shift == 'Ranná':
            shift_col = 12
        elif shift == 'Poobedna':
            shift_col = 15
        elif shift == 'Nočna':
            shift_col = 18
        else:
            return []

        parts_indices = []
        for i in range(7, 244):
            if not pd.isna(data.iloc[i, shift_col]):
                part = data.iloc[i, 3]
                if not pd.isna(part):
                    parts_indices.append(i)
        return parts_indices

    def get_project_with_parts(self, data, parts_for_shift_indices):
        project_with_parts = {}
        current_project = None
        for i in range(7, 244):
            project = data.iloc[i, 1]
            if not pd.isna(project):
                current_project = project
            if i in parts_for_shift_indices and current_project:
                part = data.iloc[i, 3]
                if current_project not in project_with_parts:
                    project_with_parts[current_project] = []
                project_with_parts[current_project].append(part)
        return project_with_parts

if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    window = ShiftApp()
    window.show()
    app.exec_()
