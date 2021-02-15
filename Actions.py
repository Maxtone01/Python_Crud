from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import sqlite3


class Insert(QDialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.btn = QPushButton()
        self.btn.setText("Registration")

        self.setWindowTitle("Adding operation")
        self.setFixedWidth(300)
        self.setFixedHeight(300)

        self.setWindowTitle("Insert operation")
        self.setFixedWidth(300)
        self.setFixedHeight(300)

        self.btn.clicked.connect(self.addperson)

        layout = QVBoxLayout()

        self.Id_enter = QLineEdit()
        self.int_check = QIntValidator()
        self.Id_enter.setValidator(self.int_check)
        self.Id_enter.setPlaceholderText("Id")
        layout.addWidget(self.Id_enter)

        self.name_enter = QLineEdit()
        self.name_enter.setPlaceholderText("Name")
        layout.addWidget(self.name_enter)

        layout.addWidget(self.btn)
        self.setLayout(layout)

    def addperson(self):
        _name = " "
        _Id = " "
        _name = self.name_enter.text()
        _Id = self.Id_enter.text()

        try:
            _connection = sqlite3.connect("MyDb.db")
            _cursor = self._connection.cursor()
            _cursor.execute("INSERT INTO person (Id, login) VALUES (?, ?)"
                            , (_Id, _name))
            _connection.commit()
            _cursor.close()
            _connection.close()
            QMessageBox.information(QMessageBox(), "Succesful", "Succesfuly added a person")
            self.close()
        except Exception:
            QMessageBox.warning(QMessageBox(), "Eror", "Could not add this person")


class Search(QDialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.btn = QPushButton()
        self.btn.setText("Search")

        self.setWindowTitle("User search")
        self.setFixedWidth(300)
        self.setFixedHeight(100)
        self.btn.clicked.connect(self.search_data)
        layout = QVBoxLayout()

        self.search = QLineEdit()
        self.int_check = QIntValidator()
        self.search.setValidator(self.int_check)
        self.search.setPlaceholderText("Name")
        layout.addWidget(self.search)
        layout.addWidget(self.btn)
        self.setLayout(layout)

    def search_data(self):
        _search = " "
        _search = self.search.text()
        try:
            _connection = sqlite3.connect("MyDb.db")
            _cursor = _connection.cursor()
            results = _cursor.execute("SELECT * FROM person WHERE Id=" + str(_search))
            row = results.fetchone()
            searchres = "Name: " + str(row[0]) + '\n' + "Id: " + str(row[1]) + '\n' + "Quantity" + str(row[2]) + '\n'
            QMessageBox.information(QMessageBox(), 'Successful ', searchres)
            _connection.commit()
            _cursor.close()
            _connection.close()
        except:
            QMessageBox.warning(QMessageBox(), "Eror", "Could not find this person")


class DropTable(QWidget):
    def droptable(self):
        reply = QMessageBox.question(self, 'Message', "You trying to delete your database, it means that "
                                                      "all your values will be deleted. "
                                                      "Do you want to do it anyway?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                _connection = sqlite3.connect("MyDb.db")
                _cursor = _connection.cursor()
                _cursor.execute("DELETE FROM person")
                QMessageBox.information(QMessageBox(), 'Successful', 'All data was deleted successful.')
                _connection.commit()
                _cursor.close()
            except:
                pass
        elif reply == QMessageBox.No:
                QMessageBox.information(QMessageBox(), 'Exit', 'Exit.')