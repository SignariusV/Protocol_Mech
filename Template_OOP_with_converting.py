# .\pyuic6  C:\Users\Vertaev\PycharmProjects\qt6\mydesign.ui -o C:\Users\Vertaev\PycharmProjects\qt6\MainWindow.py

import sys
from PyQt6 import QtWidgets, uic

from MainWindow import Ui_MainWindow


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, *args, obj=None, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.setupUi(self)

        self.pushButton.clicked.connect(self.func)

    def func(self):
        print('hehh')


app = QtWidgets.QApplication(sys.argv)

window = MainWindow()
window.show()
app.exec()