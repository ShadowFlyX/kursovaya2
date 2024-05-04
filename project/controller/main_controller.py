import sys
import os

import PySide6.QtCore
import PySide6.QtWidgets

project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../"))   # Получаем путь к каталогу проекта

sys.path.append(project_dir)                                            # Добавляем каталог проекта в PYTHONPATH
import PySide6
from PySide6.QtWidgets import QApplication, QMainWindow
from view.main_view import Ui_MainWindow
import sqlalchemy as db
from sqlalchemy.orm import sessionmaker
from models.schedule_model import ScheduleModel
from datetime import time
import openpyxl
from openpyxl.styles import (
                        PatternFill, Border, Side, 
                        Alignment, Font, GradientFill
                        )


def run_application():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())



SERVER = 'DESKTOP-HU5SQ7D\SQLEXPRESS'
DATABASE = 'TimeTable'

conntection_string = f"mssql+pyodbc://{SERVER}/{DATABASE}?trusted_connection=yes&trustservercertificate=yes&driver=ODBC+Driver+18+for+SQL+Server"
engine = db.create_engine(conntection_string)                          # Создание ядра подключения для дальнейшей работы с ним  

dirname = os.path.dirname(PySide6.__file__)                            # TODO: дописать коментарий здесь с описанием происходящего

plugin_path = os.path.join(dirname, "plugins", "platforms")
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path


Session = sessionmaker(bind=engine)                                    # Создаем сессию, чтобы

class Controller:
    def __init__(self, view: Ui_MainWindow):
        self.view = view
        self.session = Session()
        self.model = ScheduleModel()
        self.view.pushButton.clicked.connect(self.generate_file)
        self.view.comboBox.addItems(self.get_faculty_data())

    def get_faculty_data(self):
        return self.model.get_faculties(self.session)


    def generate_file(self):
        faculty = self.view.comboBox.currentText() 
        
        if faculty is None:
            error = PySide6.QtWidgets.QErrorMessage()
            error.showMessage("Выберите факультет!")
            error.setWindowTitle("Error!")
            error.exec()
            return
        
        faculty_groups = self.model.get_faculty_groups(self.session, faculty)
        semester = int(self.view.comboBox1.currentText())
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(faculty_groups) + 2)
        cell = ws.cell(row = 1, column = 1)
        cell.value = faculty
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font("Times New Roman", 20, bold=True)
    
        study_time = self.model.get_study_time(self.session, faculty, 1)

        for day in range(7):                     # 6 - кол-во учебных дней в неделе
            for row, time in enumerate(study_time, 1):
                if day == 0:
                    cell = ws.cell(row=day*7+row+2, column=2)
                else:
                    cell = ws.cell(row=day*7+row, column=2)
                cell.value = time.strftime("%#H:%M")
                cell.alignment = Alignment(vertical="center", text_rotation=90)
                cell.font = Font("Times New Roman", 10, italic=True)

        for column, groups in enumerate(faculty_groups, 3):
            group, _ = groups
            
            cell = ws.cell(row = 2, column = column)
            cell.value = group
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font("Times New Roman", 18, bold=True)
            ws.column_dimensions[openpyxl.utils.cell.get_column_letter(column)].width = 50
            
            week_schedule = self.model.get_schedule_for_week(self.session, group, semester)
            
            for row in range(60):
                ws.row_dimensions[row].height = 60
            
            for row, data in enumerate(week_schedule, 3):
                cell = ws.cell(row = row, column = column)
                cell_data = f"{data[3]}\n{data[4]}\n{data[2]}"
                cell.value = cell_data
                cell.font = Font("Times New Roman", 10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        wb.save("additional/test.xlsx")

    @staticmethod
    def process_groups_data() -> dict: ...

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controller = Controller(self.ui)

if __name__ == "__main__":
    run_application()