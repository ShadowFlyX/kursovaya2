from PySide6 import QtCore, QtGui, QtWidgets
import pandas as pd

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        MainWindow.setAcceptDrops(True)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("project/img/main_icon-transformed.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(20, 70, 191, 22))
        self.comboBox.setCurrentText("")
        self.comboBox.setPlaceholderText("")
        self.comboBox.setDuplicatesEnabled(False)
        self.comboBox.setObjectName("facultyComboBox")
        self.comboBox_2 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_2.setGeometry(QtCore.QRect(260, 70, 191, 22))
        self.comboBox_2.setObjectName("groupsComboBox")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(70, 460, 111, 24))
        self.pushButton.setObjectName("createFile")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Main Frame"))
        self.pushButton.setText(_translate("MainWindow", "Create File"))


    def create_connection(self):
        #     db.execute("SELECT * FROM TimeTable")
        #     print(db.fetchall())
        ...
    

    def get_data(self, connection):
        data = pd.read_sql_query("SELECT SGroup, Kurs, Teacher, ClassRoom, Discipline \
                           FROM TimeTable\
                           WHERE Faculty='Факультет математики и технологий программирования' and Kurs=1\
                           ORDER BY SGroup", connection)
        # connection.execute("SELECT SGroup, Kurs, Teacher, ClassRoom, Discipline \
        #                    FROM TimeTable\
        #                    WHERE Faculty='Факультет математики и технологий программирования' and Kurs=1\
        #                    ORDER BY SGroup")
        return data
    
    def process_data(self, data):
        p_data = {}

        for group, *other in data:
            p_data.setdefault(group, []).append(tuple(other))

        return p_data
    

    # def generate_file(self):
    #     with MSSQLConnection(SERVER, DATABASE) as db:
    #         data = self.get_data(db)
    #         data.head()
         
    #         # with pd.ExcelWriter("test.xlsx", mode='w') as writer:
    #         #     for line in data:
    #         #         df = pd.DataFrame(line)
    #         #         df
    #         #         # df.to_excel(writer, sheet_name="test")
        

class MSSQLConnection:

    def __init__(self, server, dbname):
        self.conn = None
        self.cursor = None
        self.server = server
        self.dbname = dbname

    def __enter__(self):
        import pyodbc
        import sqlalchemy as db



        connectionString = f'''
            DRIVER={{ODBC Driver 18 for SQL Server}};
            SERVER={self.server};DATABASE={self.dbname};
            TRUSTED_CONNECTION=YES;
            TrustServerCertificate=yes;
        '''

        
        try:
            self.conn = engine.connect()
            print("Connected successfully!")
        except Exception as e:
            print("Error:", e)
            self.conn = None
        meta = db.MetaData()
        
        
        # self.conn = pyodbc.connect(connectionString)
        # self.cursor = self.conn.cursor()
        # return self.cursor
        exit()

    def __exit__(self, exc_type, exc_value, traceback):
        if self.cursor and self.conn:
            print("Закрытие соединения")
            self.cursor.close()
            self.conn.close()


if __name__ == "__main__":
    import sys


    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    # ui.generate_file()
    MainWindow.show()
    sys.exit(app.exec())