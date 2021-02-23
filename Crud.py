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
        self._table = QTableWidget()
        self.setCentralWidget(self._table)

        self._table.setAlternatingRowColors(True)
        # Setting a number of columns
        self._table.setColumnCount(4)
        self._table.horizontalHeader().setCascadingSectionResizes(False)
        self._table.horizontalHeader().setSortIndicatorShown(False)
        self._table.horizontalHeader().setStretchLastSection(True)

        self._table.verticalHeader().setVisible(False)
        self._table.verticalHeader().setCascadingSectionResizes(False)
        self._table.verticalHeader().setStretchLastSection(False)
        self._table.setHorizontalHeaderLabels(("Id", "Teacher", "Quantity", "Grade"))

    def actions(self):
        _file_menu = self.menuBar().addMenu("&File")

        _toolbar = QToolBar()
        _toolbar.setMovable(False)
        self.addToolBar(_toolbar)

        _statusbar = QStatusBar()
        self.setStatusBar(_statusbar)

        btn_add = QAction(QIcon("./images/plus.png"), "Add user", self)
        btn_add.triggered.connect(self.insert)
        btn_add.setStatusTip("Add user")
        _toolbar.addAction(btn_add)

        btn_refresh = QAction(QIcon("./images/refresh.png"), "Refresh", self)
        btn_refresh.triggered.connect(self.refresh)
        btn_refresh.setStatusTip("Refresh")
        _toolbar.addAction(btn_refresh)

        btn_search = QAction(QIcon("./images/lupa.png"), "Search", self)
        btn_search.triggered.connect(self.search)
        btn_search.setStatusTip("Search user")
        _toolbar.addAction(btn_search)

        btn_delete = QAction(QIcon("./images/trash.png"), "Delete", self)
        btn_delete.triggered.connect(self.deleteuser)
        btn_delete.setStatusTip("Delete user")
        _toolbar.addAction(btn_delete)

        openfile = QAction(QIcon("./images/file.png"), 'Open file', self)
        openfile.triggered.connect(self.openfile)
        openfile.setStatusTip("Open file")
        _toolbar.addAction(openfile)

        btn_drop = QAction(QIcon("./images/table_delete.png"), "Delete", self)
        btn_drop.triggered.connect(self.drop_table)
        btn_drop.setStatusTip("Delete user")
        _toolbar.addAction(btn_drop)

        add_action = QAction(QIcon("./images/plus.png"), "Add person", self)
        add_action.triggered.connect(self.insert)
        _file_menu.addAction(add_action)

        search_action = QAction(QIcon("./images/lupa.png"), "Search person", self)
        search_action.triggered.connect(self.search)
        _file_menu.addAction(search_action)

        delete_action = QAction(QIcon("./images/trash.png"), "Delete user", self)
        delete_action.triggered.connect(self.deleteuser)
        _file_menu.addAction(delete_action)

        _openFile = QAction(QIcon("./images/file.png"), 'Open file', self)
        _openFile.triggered.connect(self.openfile)
        _file_menu.addAction(_openFile)

        _dropTable = QAction(QIcon("./images/table_delete.png"), 'Drop table', self)
        _dropTable.triggered.connect(self.drop_table)
        _file_menu.addAction(_dropTable)

    def refresh(self):
        _connection = sqlite3.connect("MyDb.db")
        _query = "SELECT * FROM person"
        _result = _connection.execute(_query)
        self._table.setRowCount(0)
        for _row_number, _row_data in enumerate(_result):
            self._table.insertRow(_row_number)
            for _column_number, _data in enumerate(_row_data):
                self._table.setItem(_row_number, _column_number, QTableWidgetItem(str(_data)))
        _connection.close()

    def deleteuser(self):
        _data = self._table.model().data(self._table.currentIndex())
        try:
            _connection = sqlite3.connect("MyDb.db")
            _cursor = _connection.cursor()
            _cursor.execute("DELETE FROM person WHERE Id = %s" % _data)
            _connection.commit()
            _cursor.close()
            _connection.close()
            QMessageBox.information(QMessageBox(), "Successful", "Succesfuly deleted operation")
            self.refresh()
        except Exception:
            QMessageBox.warning(QMessageBox(), "Error", "Could not delete this user. Check, if you choose table,"
                                                        " and check, did you choose an Id cell.")

    def openfile(self):
        _fname = QFileDialog.getOpenFileName(self, 'Open file', '/home')[0]
        if _fname.find('xlsx') != -1:
            book = openpyxl.open(_fname, read_only=True)
            s = book.sheetnames[0]
            sheet = book[s]
            rows = sheet.max_row
            for row in range(4, rows + 1):
                _teacher = sheet[row][0].value
                _quantity = sheet[row][1].value
                _grade = sheet[row][2].value
                _connection = sqlite3.connect("MyDb.db")
                _cursor = _connection.cursor()

                _cursor.execute("SELECT * FROM person WHERE Teacher=?",
                                (_teacher,))

                if _cursor.fetchone():
                    continue

                _cursor.execute("INSERT INTO person (Teacher, Quantity, Grade) VALUES (?, ?, ?)",
                                (_teacher, _quantity, _grade))
                _connection.commit()
                _cursor.close()
                _connection.close()

        elif _fname.find('xlsx') == -1:
            QMessageBox.warning(QMessageBox(), "Error", "There was an error while opening file. \n"
                                                        "Did you chose none excel file?")

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
