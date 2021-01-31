import sys
from Actions import *


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self._connection = sqlite3.connect("MyDb.db")
        self.cursor = self._connection.cursor()
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS person(
                        Id INTEGER,
                        login Text
                        )""")
        self.cursor.close()

        self.setWindowTitle("PyQt5 CRUD")
        self.setMinimumSize(800, 600)
        self.create_table()
        self.actions()

    def create_table(self):
        self.table = QTableWidget()
        self.setCentralWidget(self.table)

        self.table.setAlternatingRowColors(True)
        self.table.setColumnCount(2)
        self.table.horizontalHeader().setCascadingSectionResizes(False)
        self.table.horizontalHeader().setSortIndicatorShown(False)
        self.table.horizontalHeader().setStretchLastSection(True)

        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setCascadingSectionResizes(False)
        self.table.verticalHeader().setStretchLastSection(False)
        self.table.setHorizontalHeaderLabels(("Id", "Name"))

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

        add_action = QAction(QIcon("./images/plus.png"), "Add person", self)
        add_action.triggered.connect(self.insert)
        file_menu.addAction(add_action)

        search_action = QAction(QIcon("./images/lupa.png"), "Search person", self)
        search_action.triggered.connect(self.search)
        file_menu.addAction(search_action)

        delete_action = QAction(QIcon("./images/trash.png"), "Delete user", self)
        delete_action.triggered.connect(self.deleteuser)
        file_menu.addAction(delete_action)
        self.refresh()

    def refresh(self):
        self.connection = sqlite3.connect("MyDb.db")
        query = "SELECT * FROM person"
        result = self.connection.execute(query)
        self.table.setRowCount(0)
        for row_number, row_data in enumerate(result):
            self.table.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                self.table.setItem(row_number, column_number, QTableWidgetItem(str(data)))
        self.connection.close()

    def deleteuser(self):
        data = self.table.model().data(self.table.currentIndex())
        try:
            self._connection = sqlite3.connect("MyDb.db")
            self.cursor = self._connection.cursor()
            self.cursor.execute("DELETE FROM person WHERE Id = " + data)
            self._connection.commit()
            self.cursor.close()
            self._connection.close()
            QMessageBox.information(QMessageBox(), "Successful", "Succesfuly deleted operation")
            self.refresh()
        except:
            QMessageBox.warning(QMessageBox(), "Error", "Could not delete this user. You need to delete"
                                                        " user by click on Id cell.")

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
