import sys
import os

import PySide6.QtCore
import PySide6.QtWidgets
# Получаем путь к каталогу проекта
project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../"))

# Добавляем каталог проекта в PYTHONPATH
sys.path.append(project_dir)
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


engine = db.create_engine(f"mssql+pyodbc://{SERVER}/{DATABASE}?trusted_connection=yes&trustservercertificate=yes&driver=ODBC+Driver+18+for+SQL+Server")

dirname = os.path.dirname(PySide6.__file__)

plugin_path = os.path.join(dirname, "plugins", "platforms")
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path

Session = sessionmaker(bind=engine)

class Controller:
    def __init__(self, view: Ui_MainWindow):
        self.view = view
        self.session = Session()
        self.model = ScheduleModel(self.session)
        self.view.pushButton.clicked.connect(self.generate_file)
        self.view.comboBox.addItems(self.get_faculty_data())

    def get_faculty_data(self):
        return self.model.get_faculties()


    def generate_file(self):
        # faculty = self.view.comboBox.currentText() 
        faculty = None
        if faculty is None:
            error = PySide6.QtWidgets.QErrorMessage()
            error.showMessage("Выберите факультет!")
            error.setWindowTitle("Error!")
            error.exec()
            return
        data = self.model.get_groups()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controller = Controller(self.ui)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())