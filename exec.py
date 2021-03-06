import sys
import shop_tool as u1
import conf as u2
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui, QtWidgets

class SecondWindow(QMainWindow):
    def __init__(self, parent=None):
        super(SecondWindow, self).__init__(parent)
        self.ui = u2.Ui_childWidget()
        self.ui.setupUi(self)

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.ui = u1.Ui_Form()
        self.ui.setupUi(self)

    def slot1(self):
        win2.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    u1.init_db()
    win.setWindowIcon(QtGui.QIcon(":/ico/my.ico"))
    win.show()
    win2 = SecondWindow()
    sys.exit(app.exec_())