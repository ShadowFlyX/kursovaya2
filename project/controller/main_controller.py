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
        for group, _ in faculty_groups:
            week_schedule = self.model.get_schedule_for_week(self.session, group, semester)
            print(week_schedule)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controller = Controller(self.ui)

if __name__ == "__main__":
    run_application()