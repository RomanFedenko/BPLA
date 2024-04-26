from MainWindow import MainWindow
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel
import sys


def show_main(self):
    aunt.hide()
    window.show()

def show_spravka():
    spravka.show()

class Auntification(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('auntification.ui', self)
        self.pushButtonMake.clicked.connect(show_main)
        self.pushButtonSpravka.clicked.connect(show_spravka)
        self.pushButtonEN.clicked.connect(self.change_on_en)
        self.pushButtonRU.clicked.connect(self.change_on_ru)

    def change_on_en(self):
        self.labelSozdat.setText("Create your account")
        self.pushButtonSpravka.setText("Spravka")
        self.pushButtonMake.setText("Make")
        self.label_2Name.setText("Name")
        self.label_3Password.setText("Password")

    def change_on_ru(self):
        self.labelSozdat.setText("Создать ваш аккаунт")
        self.pushButtonSpravka.setText("Справка")
        self.pushButtonMake.setText("Создать")
        self.label_2Name.setText("Имя")
        self.label_3Password.setText("Пароль")

class Spravka(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('spravka.ui',self)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    aunt = Auntification()
    aunt.show()
    window = MainWindow()
    window.hide()
    spravka = Spravka()
    spravka.hide()
    sys.exit(app.exec_())

