import sys, mydb
from xlsxwriter import Workbook
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *


class Window(QWidget):
    
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Database Data Extractor")
        self.setGeometry(300, 200, 800, 600)

        main_layout = QVBoxLayout()
        self.setLayout(main_layout)
        group_box = QGroupBox("Login to Database")
        grid_layout1 = QGridLayout()
        group_box.setLayout(grid_layout1)
        main_layout.addWidget(group_box)
        self.hn_label = QLabel("Hostname:")
        self.hn_line_edit = QLineEdit()
        self.un_label = QLabel("Username:")
        self.un_line_edit = QLineEdit()
        self.pw_label = QLabel("Password:")
        self.pw_line_edit = QLineEdit()
        pw = self.pw_line_edit
        pw.setEchoMode(pw.EchoMode.Password)
        self.login_n_logout_btn = QPushButton("Login")
        self.login_n_logout_btn.clicked.connect(self.login_n_logout)
        
        # grid
        grid_layout1.addWidget(self.hn_label, 0, 0)
        grid_layout1.addWidget(self.hn_line_edit, 0, 1)
        grid_layout1.addWidget(self.un_label, 1, 0)
        grid_layout1.addWidget(self.un_line_edit, 1, 1)
        grid_layout1.addWidget(self.pw_label, 2, 0)
        grid_layout1.addWidget(self.pw_line_edit, 2, 1)
        grid_layout1.addWidget(self.login_n_logout_btn, 3, 1)

        # table widget
        self.table_widget = QTableWidget()
        main_layout.addWidget(self.table_widget)

        grid_layout2 = QGridLayout()
        main_layout.addLayout(grid_layout2)
        self.db_list_label = QLabel("Database(s):")
        self.db_combo_box = QComboBox()
        self.db_combo_box.currentTextChanged.connect(self.display_tables)
        self.label1 = QLabel("  ")
        self.table_list_label = QLabel("Table(s):")
        self.table_combo_box = QComboBox()
        self.show_records_btn = QPushButton("Show Records")
        self.show_records_btn.clicked.connect(self.show_records)

        grid_layout2.addWidget(self.db_list_label, 0, 0)
        grid_layout2.addWidget(self.db_combo_box, 0, 1)        
        grid_layout2.addWidget(self.table_list_label, 0, 2)
        grid_layout2.addWidget(self.table_combo_box, 0, 3)
        grid_layout2.addWidget(self.show_records_btn, 0, 4)

        self.dl_btn = QPushButton("Download Excel")
        self.dl_btn.clicked.connect(self.download_records)
        main_layout.addWidget(self.dl_btn)
        
        self.db_combo_box.setDisabled(True)
        self.table_combo_box.setDisabled(True)
        self.show_records_btn.setDisabled(True)
        self.dl_btn.setDisabled(True)

    def login_n_logout(self):
        self.hn = self.hn_line_edit.text()
        self.un = self.un_line_edit.text()
        self.pw = self.pw_line_edit.text()
        
        if self.login_n_logout_btn.text() != 'Logout':
            if '' in (self.hn, self.un, self.pw):
                self.msg_box("Don't leave input field(s) empty.")
            else:
                records = mydb.db_login(self.hn, self.un, self.pw)

                if type(records) != list:
                    self.msg_box(records)
                else:
                    self.hn_line_edit.setDisabled(True)
                    self.un_line_edit.setDisabled(True)
                    self.pw_line_edit.setDisabled(True)
                    self.db_combo_box.setDisabled(False)
                    self.table_combo_box.setDisabled(False)
                    self.show_records_btn.setDisabled(False)

                    self.login_n_logout_btn.setText('Logout')

                    for db in records:
                        self.db_combo_box.addItem(db[0])

                    self.msg_box('Login Successful!')
        else:
            self.hn_line_edit.setDisabled(False)
            self.un_line_edit.setDisabled(False)
            self.pw_line_edit.setDisabled(False)
            self.db_combo_box.setDisabled(True)
            self.table_combo_box.setDisabled(True)
            self.show_records_btn.setDisabled(True)
            self.dl_btn.setDisabled(True)

            self.login_n_logout_btn.setText('Login')
            self.db_combo_box.clear()
            self.table_combo_box.clear()

            row_count = self.table_widget.rowCount()
            column_count = self.table_widget.columnCount()

            self.table_widget.clear()

            # clear row and column data
            while row_count >= 0:
                self.table_widget.removeRow(row_count)
                row_count -= 1

            while column_count >= 0:
                self.table_widget.removeColumn(column_count)
                column_count -= 1

    def msg_box(self, message):
        msg = QMessageBox()
        msg.setText(message)
        msg.exec()

    def display_tables(self):
        db_name = self.db_combo_box.currentText()

        if db_name != '':
            self.table_combo_box.clear()

            tables = mydb.get_tables(self.hn, self.un, self.pw, db_name)

            for table in tables:
                self.table_combo_box.addItem(table[0])

    def show_records(self):
        self.dl_btn.setDisabled(False)

        self.db_name = self.db_combo_box.currentText()
        self.table_name = self.table_combo_box.currentText()

        records, column_names = mydb.get_records(self.hn, self.un, self.pw, self.db_name, self.table_name)
        
        self.table_widget.setRowCount(len(records))
        self.table_widget.setColumnCount(len(records[0]))

        for row_index, row_data in enumerate(records):
            for column_index, column_data in enumerate(row_data):
                
                self.table_widget.setHorizontalHeaderItem(column_index, QTableWidgetItem(str(column_names[column_index][0]))) # set column names
                self.table_widget.setItem(row_index, column_index, QTableWidgetItem(str(column_data)))
                
                if row_index % 2 != 0: # set row background color in even number
                    self.table_widget.item(row_index, column_index).setBackground(QColor('#66ff33'))

    def download_records(self):
        records, column_names = mydb.get_records(self.hn, self.un, self.pw, self.db_name, self.table_name)
        
        file_name = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel File (*.xlsx)')

        if file_name[0] != '':
            wb = Workbook(file_name[0])
            ws = wb.add_worksheet()

            bold = wb.add_format({'bold': True})

            for row_index, row_data in enumerate(records):
                for col_index, col_data in enumerate(row_data):
                    ws.write(0, col_index, column_names[col_index][0], bold)
                    ws.write(row_index + 1, col_index, col_data)

            wb.close()

    
def main():
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    app.exec()


if __name__ == '__main__':
    sys.exit(main())
