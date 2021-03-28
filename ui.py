import sys
from parsing import run
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog


class UiMainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(UiMainWindow, self).__init__()
        self.start_dir = None
        self.result_dur = None
        self.central_widget = QtWidgets.QWidget(self)
        self.open_start_dir_button = QtWidgets.QPushButton(self.central_widget)
        self.open_result_dir_button = QtWidgets.QPushButton(self.central_widget)
        self.start_text_edit = QtWidgets.QLineEdit(self.central_widget)
        self.search_label = QtWidgets.QLabel(self.central_widget)
        self.result_text_edit = QtWidgets.QLineEdit(self.central_widget)
        self.result_label = QtWidgets.QLabel(self.central_widget)
        self.search_path_line = QtWidgets.QLineEdit(self.central_widget)
        self.result_path_line = QtWidgets.QLineEdit(self.central_widget)
        self.start_button = QtWidgets.QPushButton(self.central_widget)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.central_widget)
        self.quit_button = QtWidgets.QPushButton(self.central_widget)
        self.develop_by_label = QtWidgets.QLabel(self.central_widget)
        self.resize(680, 570)
        self.setup_ui()
        self.setCentralWidget(self.central_widget)
        self.retranslate_ui(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def setup_ui(self):
        self.open_start_dir_button.setGeometry(QtCore.QRect(10, 110, 141, 32))
        self.open_start_dir_button.clicked.connect(self.get_start_directory)
        self.search_path_line.setEnabled(False)
        self.search_path_line.setGeometry(QtCore.QRect(160, 110, 501, 31))

        self.open_result_dir_button.setGeometry(QtCore.QRect(10, 160, 141, 32))
        self.open_result_dir_button.clicked.connect(self.get_result_directory)
        self.result_path_line.setEnabled(False)
        self.result_path_line.setGeometry(QtCore.QRect(160, 160, 501, 31))

        self.search_label.setGeometry(QtCore.QRect(20, 10, 91, 31))
        self.start_text_edit.setEnabled(True)
        self.start_text_edit.setGeometry(QtCore.QRect(110, 10, 551, 31))

        self.result_label.setGeometry(QtCore.QRect(20, 60, 91, 31))
        self.result_text_edit.setEnabled(True)
        self.result_text_edit.setGeometry(QtCore.QRect(110, 60, 551, 31))

        self.start_button.setGeometry(QtCore.QRect(370, 520, 141, 32))
        self.start_button.clicked.connect(self.start)

        self.plainTextEdit.setGeometry(QtCore.QRect(20, 210, 641, 301))

        self.quit_button.setGeometry(QtCore.QRect(520, 520, 141, 32))
        self.quit_button.clicked.connect(self.quit)

        self.develop_by_label.setGeometry(QtCore.QRect(20, 530, 311, 16))

    def retranslate_ui(self, main_window):
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("MainWindow", "Парсинг Excel документов"))
        self.open_start_dir_button.setText(_translate("MainWindow", "Где меняем"))
        self.open_result_dir_button.setText(_translate("MainWindow", "Куда сохраняем"))
        self.search_label.setText(_translate("MainWindow", "Найти текст:"))
        self.result_label.setText(_translate("MainWindow", "Заменить на:"))
        self.start_button.setText(_translate("MainWindow", "Старт"))
        self.quit_button.setText(_translate("MainWindow", "Выход"))
        self.develop_by_label.setText(_translate("MainWindow", "Разработчик Голов Д.Е. ООО \"Телеком Сервис\""))

    def get_start_directory(self):
        dir_list = QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        self.search_path_line.setText(dir_list)
        self.start_dir = dir_list

    def get_result_directory(self):
        dir_list = QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        self.result_path_line.setText(dir_list)
        self.result_dur = dir_list

    def start(self):
        self.plainTextEdit.appendHtml('<b style="color: blue;">Начинаю парсить...</b>')
        run(
            start_directory=self.start_dir,
            result_directory=self.result_dur,
            text_to_replace=self.start_text_edit.text(),
            result_text=self.result_text_edit.text(),
            message_screen=self.plainTextEdit
        )
        self.plainTextEdit.appendHtml('<b style="color: blue;">Успех на!</b><br>')

    def quit(self):
        sys.exit(app.exec_())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = UiMainWindow()
    ex.show()
    sys.exit(app.exec_())
