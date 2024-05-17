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
from openpyxl.styles import (Border, Side, Alignment, Font)


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


Session = sessionmaker(bind=engine)                                   

class Controller:
    COUNT_STUDY_DAYS = 6


    def __init__(self, view: Ui_MainWindow):
        self.view = view
        self.session = Session()
        self.model = ScheduleModel()
        self.view.pushButton.clicked.connect(self.generate_file)
        self.view.comboBox.addItems(self.get_faculty_data())

    def get_faculty_data(self):
        return self.model.get_faculties(self.session)

    @staticmethod
    def showError(title="Error", message="Message!"):
        error = PySide6.QtWidgets.QErrorMessage()
        error.showMessage(message)
        error.setWindowTitle(title)
        error.exec()

    @staticmethod
    def get_lesson_string(data):
        return f"{data.discipline}\n{data.teacher}\n{data.classroom}"

    def generate_file(self):
        
        #------------------------------------------------------------------------------------------------------------------
        faculty = self.view.comboBox.currentText() 
        
        if faculty is None:
            self.showError("Error", "Select faculty!")
            return
        #------------------------------------------------------------------------------------------------------------------
        
        faculty_dict = self.model.get_faculty_groups(self.session, faculty)
        # for course, groups in faculty_dict.items():
        #     self.sort_groups(groups, self.model.get_all_groups_schedule_by_course(self.session, faculty, semester, course))
        faculty_groups = [group for course_groups in faculty_dict.values() for group in course_groups]
        semester = int(self.view.comboBox1.currentText())
        
        #------------------------------------------------------------------------------------------------------------------
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        default_cell_aligment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        default_cell_font = Font("Times New Roman", 10)
        default_day_aligment = Alignment(vertical="center", horizontal="center", text_rotation=90)
        default_day_font = Font("Times New Roman", 14, bold=True)
        default_header_font = Font("Times New Roman", 20, bold=True)

        default_border = Border(left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin'))
        
        #------------------------------------------------------------------------------------------------------------------
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(faculty_groups) * 2 + 2)
        self.format_data_and_set_value(ws, 1, 1, faculty, default_cell_aligment, default_header_font, default_border)
        
        ws.merge_cells(start_row=2, start_column=1, end_row=3, end_column=2)
        self.format_data_and_set_value(ws, 2, 2, "", border=default_border)
        start_column = 3 
        for course, groups in faculty_dict.items():
            ws.merge_cells(start_row=2, start_column=start_column, end_row=2, end_column=start_column+len(groups)*2-1)
            self.format_data_and_set_value(ws, 2, start_column, f"{course} курс", default_cell_aligment, default_header_font, default_border)
            start_column += len(groups)*2
        #------------------------------------------------------------------------------------------------------------------
  

        study_time = self.model.get_study_time(self.session, faculty, semester)
        
        if not study_time:
            self.showError("Error", "No data for this faculty and semester")
            return
        
        week_days = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
        count_lessons = len(study_time)

        for day in range(self.COUNT_STUDY_DAYS):                     # 6 - кол-во учебных дней в неделе
            row_start = day * count_lessons * 2 + 4
            row_end = row_start + count_lessons * 2 - 1 
            
            ws.merge_cells(start_row=row_start, start_column=1, end_row=row_end, end_column=1)
            
            self.format_data_and_set_value(ws, row_start, 1, week_days[day], alignment=default_day_aligment, font=default_day_font, border=default_border)

            for row, time in enumerate(study_time):
                time_start_row = row_start + row*2
                ws.merge_cells(start_row=time_start_row, start_column=2, end_row=time_start_row+1, end_column=2)
                self.format_data_and_set_value(ws, time_start_row, 2, time.strftime("%#H:%M"), Alignment(vertical="center", text_rotation=90), Font("Times New Roman", 10, italic=True), default_border)

        
        #------------------------------------------------------------------------------------------------------------------
        
        for row in range(count_lessons*2*self.COUNT_STUDY_DAYS+4):
            ws.row_dimensions[row].height = 40


        column = 3
        for group in faculty_groups:
            ws.merge_cells(start_row=3, start_column=column, end_row=3, end_column=column+1) 
        
            self.format_data_and_set_value(ws, 3, column, group, default_cell_aligment, default_header_font, default_border)

            ws.column_dimensions[openpyxl.utils.cell.get_column_letter(column)].width = 15
            ws.column_dimensions[openpyxl.utils.cell.get_column_letter(column+1)].width = 15
            
            week_schedule = self.model.get_schedule_for_week(self.session, group, semester)
            
            week_schedule_dict = self.process_schedule(week_schedule, study_time)

            row = 4
            for day in week_schedule_dict:
                for time, data in week_schedule_dict[day].items():  
                    if data:   
                        # Если есть подгруппа     
                        length = len(data)
                        if data[0].overunderline is not None:
                            if data[0].overunderline.lower() == "над чертой":
                                if data[0].subgroup:
                                    if data[0].subgroup == "1":
                                        self.format_cell_and_set_value(ws, row, column, data[0], default_cell_aligment, default_cell_font, default_border)
                                        if length >= 2:
                                            if data[1].overunderline is None:
                                                ws.merge_cells(start_row=row, start_column=column+1, end_row=row+1, end_column=column+1)
                                                self.format_cell_and_set_value(ws, row, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                                if length == 3:
                                                    self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                            elif data[1].overunderline.lower() == "над чертой":
                                                self.format_cell_and_set_value(ws, row, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                                if length == 3:
                                                    if data[2].subgroup is None:
                                                        ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                        self.format_cell_and_set_value(ws, row+1, column, data[2], default_cell_aligment, default_cell_font, default_border)
                                                    elif data[2].subgroup == "1":
                                                        self.format_cell_and_set_value(ws, row+1, column, data[2], default_cell_aligment, default_cell_font, default_border)
                                                    elif data[2].subgroup == "2":
                                                        self.format_cell_and_set_value(ws, row+1, column+1, data[2], default_cell_aligment, default_cell_font, default_border)
                                                elif length == 4:
                                                    self.format_cell_and_set_value(ws, row+1, column, data[2], default_cell_aligment, default_cell_font, default_border)
                                                    self.format_cell_and_set_value(ws, row+1, column+1, data[3], default_cell_aligment, default_cell_font, default_border)
                                            elif data[1].overunderline.lower() == "под чертой":
                                                if length == 2:
                                                    if data[1].subgroup is None:
                                                        ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                        self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                                    elif data[1].subgroup == "1":
                                                        self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                                    elif data[1].subgroup == "2":
                                                        self.format_cell_and_set_value(ws, row+1, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                                elif length == 3:
                                                    self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                                    self.format_cell_and_set_value(ws, row+1, column+1, data[2], default_cell_aligment, default_cell_font, default_border)
                                    if data[0].subgroup == "2":
                                        self.format_cell_and_set_value(ws, row, column+1, data[0], default_cell_aligment, default_cell_font, default_border)
                                        if length == 2:
                                            if data[1].subgroup is None:
                                                ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                            elif data[1].subgroup == "1":
                                                self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                            elif data[1].subgroup == "2":
                                                self.format_cell_and_set_value(ws, row+1, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                        elif length == 3:
                                            self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                            self.format_cell_and_set_value(ws, row+1, column+1, data[2], default_cell_aligment, default_cell_font, default_border)
                                    if data[0].subgroup == "3":
                                        self.format_cell_and_set_value(ws, row, column+1, data[0], default_cell_aligment, default_cell_font, default_border)
                                        if length == 2:
                                            if data[1].subgroup is None:
                                                ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                                self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                            elif data[1].subgroup == "1":
                                                self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                            elif data[1].subgroup == "2":
                                                self.format_cell_and_set_value(ws, row+1, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                        elif length == 3:
                                            self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                            self.format_cell_and_set_value(ws, row+1, column+1, data[2], default_cell_aligment, default_cell_font, default_border)
                                else:
                                    ws.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column+1)
                                    self.format_cell_and_set_value(ws, row, column, data[0], default_cell_aligment, default_cell_font, default_border)
                                    if length == 2:
                                        if data[1].subgroup is None:
                                            ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                            self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                        elif data[1].subgroup == "1":
                                            self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                        elif data[1].subgroup == "2":
                                            self.format_cell_and_set_value(ws, row+1, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                    elif length == 3:
                                        self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                        self.format_cell_and_set_value(ws, row+1, column+1, data[2], default_cell_aligment, default_cell_font, default_border)
                                    else:
                                        ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                        self.format_cell_and_set_value(ws, row+1, column, "", default_cell_aligment, default_cell_font, default_border)
                            elif data[0].overunderline.lower() == "под чертой":
                                if length == 1:
                                    if data[0].subgroup is None:
                                        ws.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column+1)
                                        self.format_cell_and_set_value(ws, row, column, "", default_cell_aligment, default_cell_font, default_border)
                                        ws.merge_cells(start_row=row+1, start_column=column, end_row=row+1, end_column=column+1)
                                        self.format_cell_and_set_value(ws, row+1, column, data[0], default_cell_aligment, default_cell_font, default_border)
                                    elif data[0].subgroup == "1":
                                        self.format_cell_and_set_value(ws, row+1, column, data[0], default_cell_aligment, default_cell_font, default_border)
                                    elif data[0].subgroup == "2":
                                        self.format_cell_and_set_value(ws, row+1, column+1, data[0], default_cell_aligment, default_cell_font, default_border)
                                    elif data[0].subgroup == "3":
                                        self.format_cell_and_set_value(ws, row+1, column, data[0], default_cell_aligment, default_cell_font, default_border)
                                elif length == 2:
                                    ws.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column+1)
                                    self.format_cell_and_set_value(ws, row, column, "", default_cell_aligment, default_cell_font, default_border)
                                    self.format_cell_and_set_value(ws, row+1, column, data[0], default_cell_aligment, default_cell_font, default_border)
                                    self.format_cell_and_set_value(ws, row+1, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                        else:
                            if data[0].subgroup is not None:
                                if data[0].subgroup == "1":
                                    ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column)
                                    self.format_cell_and_set_value(ws, row, column, data[0], default_cell_aligment, default_cell_font, default_border)
                                    if length == 3:
                                        self.format_cell_and_set_value(ws, row, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                        self.format_cell_and_set_value(ws, row+1, column+1, data[2], default_cell_aligment, default_cell_font, default_border)
                                    elif length == 2:
                                        if data[1].overunderline is not None:
                                            if data[1].overunderline.lower() == "над чертой":
                                                self.format_cell_and_set_value(ws, row, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                            elif data[1].overunderline.lower() == "под чертой":
                                                self.format_cell_and_set_value(ws, row+1, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                        else:
                                            ws.merge_cells(start_row=row, start_column=column+1, end_row=row+1, end_column=column+1)
                                            self.format_cell_and_set_value(ws, row, column+1, data[1], default_cell_aligment, default_cell_font, default_border)
                                    else:
                                        ws.merge_cells(start_row=row, start_column=column+1, end_row=row+1, end_column=column+1)
                                        self.format_cell_and_set_value(ws, row, column+1, "", default_cell_aligment, default_cell_font, default_border)
                                elif data[0].subgroup == "2":
                                    ws.merge_cells(start_row=row, start_column=column+1, end_row=row+1, end_column=column+1)
                                    self.format_cell_and_set_value(ws, row, column+1, data[0], default_cell_aligment, default_cell_font, default_border)
                                    ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column)
                                    self.format_cell_and_set_value(ws, row, column, "", default_cell_aligment, default_cell_font, default_border)
                                    if length == 2:
                                        self.format_cell_and_set_value(ws, row+1, column, data[1], default_cell_aligment, default_cell_font, default_border)
                                elif data[0].subgroup == "3" and length == 1:
                                    ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column+1)
                                    self.format_cell_and_set_value(ws, row, column, data[0], default_cell_aligment, default_cell_font, default_border)
                            else:
                                ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column+1)
                                self.format_cell_and_set_value(ws, row, column, data[0], default_cell_aligment, default_cell_font, default_border)
                    else:
                        ws.merge_cells(start_row=row, start_column=column, end_row=row+1, end_column=column+1)
                        self.format_cell_and_set_value(ws, row, column, "", default_cell_aligment, default_cell_font, default_border)
                    row += 2
            column += 2  

        #------------------------------------------------------------------------------------------------------------------

        wb.save("additional/test.xlsx")

    def format_cell_and_set_value(self, ws, row, column, data, alignment=None, font=None, border=None):
        cell = ws.cell(row=row, column=column)
        if data:
            cell.value = self.get_lesson_string(data)
        if alignment:
            cell.alignment = alignment
        if font:
            cell.font = font
        if border:
            cell.border = border

    def format_data_and_set_value(self, ws, row, column, data, alignment=None, font=None, border=None):
        cell = ws.cell(row=row, column=column)
        if data:
            cell.value = data
        if alignment:
            cell.alignment = alignment
        if font:
            cell.font = font
        if border:
            cell.border = border

    def sort_groups(self, groups, data):
        dct = {group: [] for group in groups}
        
        ...


    def process_schedule(self, week_schedule, study_time):
        week_schedule_dict = {}

        for day in range(1, self.COUNT_STUDY_DAYS+1):
            week_schedule_dict.setdefault(day, {})
            for time in study_time:
                week_schedule_dict[day].setdefault(time, [])

        for data in week_schedule:
            week_schedule_dict[data.dayofweek][data.timebeg].append(data)
    
        return week_schedule_dict


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controller = Controller(self.ui)

if __name__ == "__main__":
    run_application()