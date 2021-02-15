import sys
from Actions import *
import openpyxl
import asyncio

connection = sqlite3.connect("MyDb.db")
cursor = connection.cursor()
cursor.execute("""CREATE TABLE IF NOT EXISTS person(
                        ID INTEGER PRIMARY KEY,
                        Teacher TEXT,
                        Quantity TEXT,
                        Grade TEXT
                        )""")
cursor.close()


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.setWindowTitle("PyQt5 CRUD")
        self.setMinimumSize(800, 600)
        self.create_table()
        self.actions()

    def create_table(self):
        self.table = QTableWidget()
        self.setCentralWidget(self.table)

        self.table.setAlternatingRowColors(True)
        # Setting a number of columns
        self.table.setColumnCount(4)
        self.table.horizontalHeader().setCascadingSectionResizes(False)
        self.table.horizontalHeader().setSortIndicatorShown(False)
        self.table.horizontalHeader().setStretchLastSection(True)

        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setCascadingSectionResizes(False)
        self.table.verticalHeader().setStretchLastSection(False)
        self.table.setHorizontalHeaderLabels(("Id", "Teacher", "Quantity", "Grade"))

    def actions(self):
        file_menu = self.menuBar().addMenu("&File")

        toolbar = QToolBar()
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        statusbar = QStatusBar()
        self.setStatusBar(statusbar)

        btn_add = QAction(QIcon("./images/plus.png"), "Add user", self)
        btn_add.triggered.connect(self.insert)
        btn_add.setStatusTip("Add user")
        toolbar.addAction(btn_add)

        btn_refresh = QAction(QIcon("./images/refresh.png"), "Refresh", self)
        btn_refresh.triggered.connect(self.refresh)
        btn_refresh.setStatusTip("Refresh")
        toolbar.addAction(btn_refresh)

        btn_search = QAction(QIcon("./images/lupa.png"), "Search", self)
        btn_search.triggered.connect(self.search)
        btn_search.setStatusTip("Search user")
        toolbar.addAction(btn_search)

        btn_delete = QAction(QIcon("./images/trash.png"), "Delete", self)
        btn_delete.triggered.connect(self.deleteuser)
        btn_delete.setStatusTip("Delete user")
        toolbar.addAction(btn_delete)

        openFile = QAction(QIcon("./images/file.png"), 'Open file', self)
        openFile.triggered.connect(self.openfile)
        openFile.setStatusTip("Open file")
        toolbar.addAction(openFile)

        btn_drop = QAction(QIcon("./images/table_delete.png"), "Delete", self)
        btn_drop.triggered.connect(self.drop_table)
        btn_drop.setStatusTip("Delete user")
        toolbar.addAction(btn_drop)

        add_action = QAction(QIcon("./images/plus.png"), "Add person", self)
        add_action.triggered.connect(self.insert)
        file_menu.addAction(add_action)

        search_action = QAction(QIcon("./images/lupa.png"), "Search person", self)
        search_action.triggered.connect(self.search)
        file_menu.addAction(search_action)

        delete_action = QAction(QIcon("./images/trash.png"), "Delete user", self)
        delete_action.triggered.connect(self.deleteuser)
        file_menu.addAction(delete_action)

        _openFile = QAction(QIcon("./images/file.png"), 'Open file', self)
        _openFile.triggered.connect(self.openfile)
        file_menu.addAction(_openFile)

        _dropTable = QAction(QIcon("./images/table_delete.png"), 'Drop table', self)
        _dropTable.triggered.connect(self.drop_table)
        file_menu.addAction(_dropTable)

    def refresh(self):
        try:
            _connection = sqlite3.connect("MyDb.db")
            _query = "SELECT * FROM person"
            _result = _connection.execute(_query)
            self.table.setRowCount(0)
            for row_number, row_data in enumerate(_result):
                self.table.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    self.table.setItem(row_number, column_number, QTableWidgetItem(str(data)))
            _connection.close()
        except:
            pass

    def deleteuser(self):
        data = self.table.model().data(self.table.currentIndex())
        try:
            _connection = sqlite3.connect("MyDb.db")
            _cursor = _connection.cursor()
            _cursor.execute("DELETE FROM person WHERE Id = %s" % data)
            _connection.commit()
            _cursor.close()
            _connection.close()
            QMessageBox.information(QMessageBox(), "Successful", "Succesfuly deleted operation")
            self.refresh()
        except:
            QMessageBox.warning(QMessageBox(), "Error", "Could not delete this user. You need to delete"
                                                        " user by click on Id cell.")

    def openfile(self):
        try:
            fname = QFileDialog.getOpenFileName(self, 'Open file', '/home')[0]
            if fname.find('xlsx') != -1:
                book = openpyxl.open(fname, read_only=True)
                s = book.sheetnames[0]
                sheet = book[s]
                rows = sheet.max_row
                for row in range(4, rows + 1):
                    _teacher = sheet[row][0].value
                    _quantity = sheet[row][1].value
                    _grade = sheet[row][2].value
                    _connection = sqlite3.connect("MyDb.db")
                    _cursor = _connection.cursor()
                    _cursor.execute("INSERT INTO person (Teacher, Quantity, Grade) VALUES (?, ?, ?)",
                                (_teacher, _quantity, _grade))
                    _connection.commit()
                    _cursor.close()
                    _connection.close()
            else:
                QMessageBox.warning(QMessageBox(), "Error", "There was an error while opening file. \n"
                                                            "Did you chose none excel file?")
        except:
            pass

    def drop_table(self):
        dlg = DropTable()
        dlg.droptable()

    def insert(self):
        dlg = Insert()
        dlg.exec_()
        self.refresh()

    def search(self):
        dlg = Search()
        dlg.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
