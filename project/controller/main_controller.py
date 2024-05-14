import sys
import os

import PySide6
from PySide6.QtWidgets import QApplication, QMainWindow
from view.main_view import Ui_MainWindow
import sqlalchemy as db
from sqlalchemy.orm import sessionmaker
from models.schedule_model import ScheduleModel
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, GradientFill

def run_application():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

SERVER = 'DESKTOP-HU5SQ7D\SQLEXPRESS'
DATABASE = 'TimeTable'

conntection_string = f"mssql+pyodbc://{SERVER}/{DATABASE}?trusted_connection=yes&trustservercertificate=yes&driver=ODBC+Driver+18+for+SQL+Server"
engine = db.create_engine(conntection_string)

dirname = os.path.dirname(PySide6.__file__)

plugin_path = os.path.join(dirname, "plugins", "platforms")
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path

Session = sessionmaker(bind=engine)

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
        faculty = self.view.comboBox.currentText()

        if faculty is None:
            self.show_error_message("Выберите факультет!")
            return

        faculty_groups = self.model.get_faculty_groups(self.session, faculty)
        semester = int(self.view.comboBox1.currentText())

        wb = openpyxl.Workbook()
        ws = wb.active

        self.merge_and_set_value(ws, 1, 1, 1, len(faculty_groups) + 2, faculty, font_size=20, bold=True)

        study_time = self.model.get_study_time(self.session, faculty, semester)

        week_days = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
        count_lessons = len(study_time)
        COUNT_STUDY_DAYS = 6

        for day in range(COUNT_STUDY_DAYS):
            row_start = day * count_lessons * 2 + 3
            row_end = row_start + count_lessons * 2 - 1

            self.merge_and_set_value(ws, row_start, 1, row_end, 1, week_days[day],
                                     vertical_align="center", horizontal_align="center", text_rotation=90, font_size=14)

            for row, time in enumerate(study_time):
                time_start_row = row_start + row * 2
                self.merge_and_set_value(ws, time_start_row, 2, time_start_row + 1, 2, time.strftime("%#H:%M"),
                                         vertical_align="center", font_size=10, italic=True)

        column = 3
        for groups in faculty_groups:
            group, _ = groups
            self.merge_and_set_value(ws, 2, column, 2, column + 1, group, vertical_align="center", horizontal_align="center", font_size=18, bold=True)
            ws.column_dimensions[openpyxl.utils.cell.get_column_letter(column)].width = 15
            ws.column_dimensions[openpyxl.utils.cell.get_column_letter(column + 1)].width = 15

            week_schedule = self.model.get_schedule_for_week(self.session, group, semester)

            for row in range(100):
                ws.row_dimensions[row].height = 40

            week_schedule_dict = {}
            for day in range(1, COUNT_STUDY_DAYS + 1):
                week_schedule_dict.setdefault(day, {})
                for time in study_time:
                    week_schedule_dict[day].setdefault(time, [])

            for data in week_schedule:
                week_schedule_dict[data.dayofweek][data.timebeg].append(data)

            row = 3
            for day in week_schedule_dict:
                for time, data in week_schedule_dict[day].items():
                    if data:
                        self.process_schedule_data(ws, row, column, data)
                    else:
                        ws.merge_cells(start_row=row, start_column=column, end_row=row + 1, end_column=column + 1)
                    row += 2
            column += 2

        wb.save("additional/test.xlsx")

    def process_schedule_data(self, ws, row, column, data):
        length = len(data)
        if data[0].overunderline is not None:
            if data[0].overunderline.lower() == "над чертой":
                self.process_above_line(ws, row, column, data, length)
            elif data[0].overunderline.lower() == "под чертой":
                self.process_below_line(ws, row, column, data, length)
        else:
            self.process_no_line(ws, row, column, data, length)

    def process_above_line(self, ws, row, column, data, length):
        if data[0].subgroup:
            if data[0].subgroup == "1":
                self.set_subgroup(ws, row, column, data, length, 1)
            if data[0].subgroup == "2":
                self.set_subgroup(ws, row, column, data, length, 2)
        else:
            self.merge_and_set_value(ws, row, column, row, column + 1, self.get_lesson_string(data[0]))
            self.set_subgroup(ws, row, column, data, length, None)

    def process_below_line(self, ws, row, column, data, length):
        self.set_subgroup(ws, row + 1, column, data, length, None)

    def process_no_line(self, ws, row, column, data, length):
        if data[0].subgroup:
            if data[0].subgroup == "1":
                self.merge_and_set_value(ws, row, column, row + 1, column, self.get_lesson_string(data[0]))
                self.set_subgroup(ws, row, column + 1, data, length, None)
            elif data[0].subgroup == "2":
                self.merge_and_set_value(ws, row, column + 1, row + 1, column + 1, self.get_lesson_string(data[0]))
                self.set_subgroup(ws, row + 1, column, data, length, None)
        else:
            self.merge_and_set_value(ws, row, column, row + 1, column + 1, self.get_lesson_string(data[0]))

    def set_subgroup(self, ws, row, column, data, length, subgroup):
        for i in range(1, length):
            if data[i].subgroup == subgroup or subgroup is None:
                ws.cell(row=row, column=column).value = self.get_lesson_string(data[i])

    def show_error_message(self, message):
        error = PySide6.QtWidgets.QErrorMessage()
        error.showMessage(message)
        error.setWindowTitle("Error!")
        error.exec()

    def format_cell(self, cell, vertical_align=None, horizontal_align=None, text_rotation=None, font_size=None, bold=False, italic=False):
        if vertical_align:
            cell.alignment = cell.alignment.copy(vertical=vertical_align)
        if horizontal_align:
            cell.alignment = cell.alignment.copy(horizontal=horizontal_align)
        if text_rotation:
            cell.alignment = cell.alignment.copy(textRotation=text_rotation)
        if font_size:
            cell.font = cell.font.copy(size=font_size)
        if bold:
            cell.font = cell.font.copy(bold=bold)
        if italic:
            cell.font = cell.font.copy(italic=italic)

    def merge_and_set_value(self, ws, start_row, start_column, end_row, end_column, value, **kwargs):
        ws.merge_cells(start_row=start_row, start_column=start_column, end_row=end_row, end_column=end_column)
        cell = ws.cell(row=start_row, column=start_column)
        cell.value = value
        self.format_cell(cell, **kwargs)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controller = Controller(self.ui)

if __name__ == "__main__":
    run_application()
