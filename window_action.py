import sys
from PyQt5.QtWidgets import (QPushButton, QApplication, QMessageBox, QDesktopWidget, QLabel,
                             QComboBox, QFileDialog,  QMainWindow,  QHBoxLayout,)
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import *
from slice import SliceXLS
from logger_config import get_logger
log = get_logger('logger_main')

StyleSheet = '''
#BlueProgressBar {
    border: 2px solid #2196F3;
    border-radius: 5px;
    background-color: #E0E0E0;
}
'''


class SliceWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()
        self.file_name = ''
        self.xls_file = None
        self.selected_sheet = ''
        self.selected_column = ''
        self.destination_dir = ''


    def initUI(self):
        # self.setGeometry(300, 300, 300, 220)
        self.resize(400, 400)
        self.setWindowTitle('Разбиенеие XLS файлов')
        self.setWindowIcon(QIcon('data/ms-excel-icon.png'))

        self.btn_source_file = QPushButton('Исходный файл', self)
        self.btn_source_file.resize(self.btn_source_file.sizeHint())
        self.btn_source_file.move(250, 50)
        self.btn_source_file.clicked.connect(self.get_source_file)

        self.btn_dist_dir = QPushButton('Куда складывать', self)
        self.btn_dist_dir.resize(self.btn_dist_dir.sizeHint())
        self.btn_dist_dir.setEnabled(False)
        self.btn_dist_dir.move(250, 80)
        self.btn_dist_dir.clicked.connect(self.get_destination)

        self.btn_start_split = QPushButton('Разбить', self)
        self.btn_start_split.resize(self.btn_start_split.sizeHint())
        self.btn_start_split.setEnabled(False)
        self.btn_start_split.move(250, 210)
        self.btn_start_split.clicked.connect(self.start_split)
        # self.btn_start_split.clicked.connect(QCoreApplication.instance().quit)

        self.label_file = QLabel(self)
        self.label_file.resize(180, 15)
        self.label_file.move(20, 50)

        self.label_dir = QLabel(self)
        self.label_dir.resize(180, 15)
        self.label_dir.move(20, 80)

        self.box_file = QHBoxLayout(self)
        self.box_file.addWidget(self.btn_source_file)
        self.box_file.addWidget(self.label_file)

        self.box_dir = QHBoxLayout(self)
        self.box_dir.addWidget(self.btn_dist_dir)
        self.box_dir.addWidget(self.label_dir)

        self.list_sheets = QComboBox(self)
        self.list_sheets.move(50, 140)
        self.list_sheets.activated[str].connect(self.select_sheet)

        self.list_column = QComboBox(self)
        self.list_column.move(50, 180)
        self.list_column.activated[str].connect(self.select_column)

        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def get_source_file(self):
        self.file_name = QFileDialog.getOpenFileName(self, 'Single File', 'C:\\', '*.xlsx *.xls')[0]
        self.btn_dist_dir.setEnabled(True)
        self.label_file.setText(self.file_name)
        self.xls_file = SliceXLS(self.file_name)
        self.combobox_sheets()

    def get_destination(self):
        self.destination_dir = QFileDialog.getExistingDirectory(None, "Select Folder")
        self.label_dir.setText(self.destination_dir)

    def combobox_sheets(self):
        self.list_sheets.clear()
        sheet_names = self.xls_file.read_sheets()
        self.list_sheets.addItems(sheet_names)

    def select_sheet(self, sheet):
        self.selected_sheet = sheet
        column_names = self.xls_file.read_data(sheet)
        self.list_column.clear()
        self.list_column.addItems(column_names)

    def select_column(self, column):
        self.selected_column = column
        self.btn_start_split.setEnabled(True)

    def start_split(self):
        if self.xls_file.split_file(self.selected_sheet, self.selected_column, self.destination_dir):
            QMessageBox.information(None, 'Успех', "Да ладно! у тебя получилось!")
        else:
            QMessageBox.information(None, 'Ошибка', "Что-то пошло не так(((")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SliceWindow()
    sys.exit(app.exec_())

