import sys
from PyQt5.QtWidgets import (QMainWindow, QTextEdit,
    QAction, QFileDialog, QApplication)
from PyQt5.QtGui import QIcon
from stat_func import *


class Example(QMainWindow):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        self.textEdit = QTextEdit()
        self.setCentralWidget(self.textEdit)
        self.statusBar()

        openFile = QAction(QIcon('open.png'), 'Open', self)
        openFile.setShortcut('Ctrl+O')
        openFile.setStatusTip('Open new File')
        openFile.triggered.connect(self.showDialog)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(openFile)

        self.setGeometry(300, 300, 350, 300)
        self.setWindowTitle('File dialog')
        self.show()

    def showDialog(self):
        file_path = QFileDialog.getOpenFileName(self, 'Open file', '/thome')[0]
        wb = load_workbook(file_path)
        sheets = wb.worksheets

        all_sheets_count = len(sheets)
        counter = 1
        self.textEdit.setText("Run!")
        print("RUN!")

        for sheet in sheets:
            self.textEdit.append("Processed sheet - %s, from %s" % (counter, all_sheets_count))
            print("Processed sheet - %s, from %s" % (counter, all_sheets_count))

            check_sheet = check_headers(sheet)

            if not check_sheet:
                prepare_sheet(sheet)

            for row_num in range(sheet.max_row, 1, -1):
                take_data(sheet, row_num)

            write_data(sheet)
            SHEET_DATA.clear()
            counter += 1

        wb.save(file_path)
        self.textEdit.append("Successfully processed")
        print("Done")

if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
