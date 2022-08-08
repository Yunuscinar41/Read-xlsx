from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5 import QtGui
import sys
import datetime
from openpyxl import load_workbook
import pymysql

class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.setGeometry(600,200,400,400)
        self.setWindowTitle("Insert Customers")
        self.initUI()

    def initUI(self):
        self.hostName = QtWidgets.QLabel(self)
        self.hostName.setText("Hostname ")
        self.hostName.setFont(QtGui.QFont("Sanserif", 15))
        self.hostName.move(50,50)

        self.hostName_ = QtWidgets.QLineEdit(self)
        self.hostName_.setFont(QtGui.QFont("Sanserif", 15))
        self.hostName_.move(50,80)

        self.username = QtWidgets.QLabel(self)
        self.username.setText("Username")
        self.username.setFont(QtGui.QFont("Sanserif", 15))
        self.username.move(220,50)

        self.username_ = QtWidgets.QLineEdit(self)
        self.username_.setFont(QtGui.QFont("Sanserif", 15))
        self.username_.move(220,80)


        self.password = QtWidgets.QLabel(self)
        self.password.setText("Password")
        self.password.setFont(QtGui.QFont("Sanserif", 15))
        self.password.move(50,130)

        self.password_ = QtWidgets.QLineEdit(self)
        self.password_.setFont(QtGui.QFont("Sanserif", 15))
        self.password_.move(50,160)

        self.database = QtWidgets.QLabel(self)
        self.database.setText("Database")
        self.database.setFont(QtGui.QFont("Sanserif", 15))
        self.database.move(220,130)

        self.database_ = QtWidgets.QLineEdit(self)
        self.database_.setFont(QtGui.QFont("Sanserif", 15))
        self.database_.move(220,160)

        self.path = QtWidgets.QLabel(self)
        self.path.setText("Path")
        self.path.setFont(QtGui.QFont("Sanserif", 15))
        self.path.move(135,200)

        self.path_ = QtWidgets.QLineEdit(self)
        self.path_.setFont(QtGui.QFont("Sanserif", 15))
        self.path_.move(135,230)

        self.push = QtWidgets.QPushButton(self)
        self.push.setText("Read!")
        self.push.clicked.connect(self.clicked)
        self.push.move(135,290)

        self.error = QtWidgets.QLabel(self)
        self.error.setText("")
        self.error.setFont(QtGui.QFont("Sanserif", 15))
        self.error.move(60,340)

    def clicked(self):
        try:
            self.do()
            self.error.setText("Read and Insert Successfull!!")
            self.update()
        except:
            self.error.setText("An exception occurred!")
            self.update()
    def update(self):
        self.error.adjustSize()

    def do(self):
        sheet = load_workbook(self.path_.text()).active

        rows = sheet.rows
        count = 0

        headers = [cell.value for cell in next(rows)]
        all_rows = []

        for row in rows :
            if(count == 173):
                break

            data = []
            for title, cell in zip(headers, row) :

                data.append(cell.value)

                if(title == "Musteri"):
                    break

            all_rows.append(data)
            count += 1


        conn = pymysql.connect(
            host = self.hostName_.text(),
            user = self.username_.text(),
            password = self.password_.text(),
            database = self.database_.text()
        )

        cursor = conn.cursor()

        sql = "INSERT INTO müsteri (`İsim`, `Cihaz Seri No`) VALUES (%s, %s)"

        for i in range(len(all_rows)):
            try:
                cursor.execute(sql,all_rows[i])
            except:
                count = 0
        conn.commit()
        conn.close()



def window():
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    sys.exit(app.exec_())

window()
