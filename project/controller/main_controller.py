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

    @staticmethod
    def get_lesson_string(data):
        return f"{data.discipline}\n{data.teacher}\n{data.classroom}"

    def generate_file(self):
        
        #------------------------------------------------------------------------------------------------------------------
        faculty = self.view.comboBox.currentText() 
        
        if faculty is None:
            error = PySide6.QtWidgets.QErrorMessage()
            error.showMessage("Выберите факультет!")
            error.setWindowTitle("Error!")
            error.exec()
            return
        #------------------------------------------------------------------------------------------------------------------
        
        faculty_groups = self.model.get_faculty_groups(self.session, faculty)
        semester = int(self.view.comboBox1.currentText())
        
        #------------------------------------------------------------------------------------------------------------------
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        #------------------------------------------------------------------------------------------------------------------
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(faculty_groups) + 2)
        cell = ws.cell(row=1, column=1)
        cell.value = faculty
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font("Times New Roman", 20, bold=True)
        #------------------------------------------------------------------------------------------------------------------
  
        study_time = self.model.get_study_time(self.session, faculty, semester)
        
        week_days = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
        count_lessons = len(study_time)
        COUNT_STUDY_DAYS = 6

        for day in range(COUNT_STUDY_DAYS):                     # 6 - кол-во учебных дней в неделе
            row_start = day * count_lessons * 2 + 3
            row_end = row_start + count_lessons * 2 - 1 
            
            ws.merge_cells(start_row=row_start, start_column=1, end_row=row_end, end_column=1)
            
            cell = ws.cell(row=row_start, column=1)
            cell.value = week_days[day]
            
            cell.alignment = Alignment(vertical="center", horizontal="center", text_rotation=90)
            cell.font = Font("Times New Roman", 14)
            
            for row, time in enumerate(study_time):
                time_start_row = row_start + row*2
                ws.merge_cells(start_row=time_start_row, start_column=2, end_row=time_start_row+1, end_column=2)
                cell = ws.cell(row=row_start+row*2, column=2)
                cell.value = time.strftime("%#H:%M")
                cell.alignment = Alignment(vertical="center", text_rotation=90)
                cell.font = Font("Times New Roman", 10, italic=True)
        
        #------------------------------------------------------------------------------------------------------------------
        
        column = 3
        for groups in faculty_groups:
            group, _ = groups
            ws.merge_cells(start_row=2, start_column=column, end_row=2, end_column=column+1) 
            
            cell = ws.cell(row = 2, column = column)

            cell.value = group
            
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font("Times New Roman", 18, bold=True)
            ws.column_dimensions[openpyxl.utils.cell.get_column_letter(column)].width = 15
            ws.column_dimensions[openpyxl.utils.cell.get_column_letter(column+1)].width = 15
            
            week_schedule = self.model.get_schedule_for_week(self.session, group, semester)
            
            for row in range(100):
                ws.row_dimensions[row].height = 40
            

            week_schedule_dict = {}

            for day in range(1, COUNT_STUDY_DAYS+1):
                week_schedule_dict.setdefault(day, {})
                for time in study_time:
                    week_schedule_dict[day].setdefault(time, [])


            for data in week_schedule:
                week_schedule_dict[data.dayofweek][data.timebeg].append(data)

            row = 3
            for day in week_schedule_dict:
                for time, data in week_schedule_dict[day].items():  
                    if data:   
                        # Если есть подгруппа     
                        length = len(data)
                        if data[0].overunderline is not None:
                            if data[0].overunderline.lower() == "над чертой":
                                if data[0].subgroup:
                                    if data[0].subgroup == "1":
                                        ws.cell(row=row, column=column).value = self.get_lesson_string(data[0])
                                        if length >= 2:
                                            if data[1].overunderline is None:
                                                ws.merge_cells(start_row=row, start_column=column+1, end_row=row+1, end_column=column+1)
                                                ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[1])
                                                if length == 3:
                                                    ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                            elif data[1].overunderline.lower() == "над чертой":
                                                ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[1])
                                                if length == 3:
                                                    if data[2].subgroup is None:
                                                        ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[2])
                                                    elif data[2].subgroup == "1":
                                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[2])
                                                    elif data[2].subgroup == "2":
                                                        ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[2])
                                                elif length == 4:
                                                    ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[2])  
                                                    ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[3])
                                            elif data[1].overunderline.lower() == "под чертой":
                                                if length == 2:
                                                    if data[1].subgroup is None:
                                                        ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                                    elif data[1].subgroup == "1":
                                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                                    elif data[1].subgroup == "2":
                                                        ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[1])
                                                elif length == 3:
                                                    ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])  
                                                    ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[2])
                                    if data[0].subgroup == "2":
                                        ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[0])
                                        if length == 2:
                                            if data[1].subgroup is None:
                                                ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                            elif data[1].subgroup == "1":
                                                ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                            elif data[1].subgroup == "2":
                                                ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[1])
                                        elif length == 3:
                                            ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])  
                                            ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[2])
                                    if data[0].subgroup == "3":
                                        ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[0])
                                        if length == 2:
                                            if data[1].subgroup is None:
                                                ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                            elif data[1].subgroup == "1":
                                                ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                            elif data[1].subgroup == "2":
                                                ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[1])
                                        elif length == 3:
                                            ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])  
                                            ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[2])
                                else:
                                    ws.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column+1)
                                    ws.cell(row=row, column=column).value = self.get_lesson_string(data[0])
                                    if length == 2:
                                        if data[1].subgroup is None:
                                            ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                            ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                        elif data[1].subgroup == "1":
                                            ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                                        elif data[1].subgroup == "2":
                                            ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[1])
                                    elif length == 3:
                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])  
                                        ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[2])
                            elif data[0].overunderline.lower() == "под чертой":
                                if length == 1:
                                    if data[0].subgroup is None:
                                        ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[0])
                                    elif data[0].subgroup == "1":
                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[0])
                                    elif data[0].subgroup == "2":
                                        ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[0])
                                    elif data[0].subgroup == "3":
                                        # Чисто для физики ©Миша
                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[0])
                                elif length == 2:
                                    ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[0])  
                                    ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[1])
                        else:
                            # Если есть подгруппа
                            if data[0].subgroup is not None:
                                if data[0].subgroup == "1":
                                    ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column)
                                    ws.cell(row=row, column=column).value = self.get_lesson_string(data[0]) 
                                    if length == 3:
                                        ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[1]) 
                                        ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[2]) 
                                    elif length == 2:
                                        if data[1].overunderline is not None:
                                            if data[1].overunderline.lower() == "над чертой":
                                                ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[1]) 
                                            elif data[1].overunderline.lower() == "под чертой":
                                                ws.cell(row=row+1, column=column+1).value = self.get_lesson_string(data[1]) 
                                        else:
                                            ws.merge_cells(start_row=row, start_column=column+1, end_row=row+1, end_column=column+1)
                                            ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[1]) 
                                elif data[0].subgroup == "2":
                                    ws.merge_cells(start_row=row, start_column=column+1, end_row=row+1, end_column=column+1)
                                    ws.cell(row=row, column=column+1).value = self.get_lesson_string(data[0]) 
                                    if length == 2:
                                        ws.cell(row=row+1, column=column).value = self.get_lesson_string(data[1])
                            else:
                                ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column+1)
                                cell = ws.cell(row=row, column=column)
                                ws.cell(row=row, column=column).value = self.get_lesson_string(data[0]) 
                    else:
                        ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column+1)
                    row += 2
            column += 2  
            '''
            cell.value = cell_data
            cell.font = Font("Times New Roman", 10)
            cell.alignment = Alignment(horizontal="center", vertical="center")'''  
        #------------------------------------------------------------------------------------------------------------------

        wb.save("additional/test.xlsx")


class Lessons:

    def __init__(self, arr: list):
        self._size = len(arr)
        self._ul = None
        self._ur = None
        self._dl = None
        self._dr = None

        self._lessons = self.process_array(arr)


    def process_array(self, arr: list):
        for index in range(self._size):
            if arr[index].overunderline is not None:
                if index+1 < self._size:
                    if arr[index].overunderline.lower() == "под чертой":
                        if arr[index+1].overunderline == arr[index].overunderline:
                            self._ul = arr[index]
                            self._ur = arr[index+1]
                
        



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controller = Controller(self.ui)

if __name__ == "__main__":
    run_application()