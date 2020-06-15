import sys, time, random, os
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtWidgets import QAbstractItemView
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine, exc, inspect, select, MetaData, \
    Column, Integer, String
from pprint import pprint
from ui_search import Ui_Form


Base = declarative_base()
engine = create_engine('sqlite:///dpr.db')
Base.metadata.create_all(engine)
# Create a session to handle updates.
Session = sessionmaker(bind=engine)
session = Session()
metadata = MetaData()
metadata.reflect(bind=engine)
inspector = inspect(engine)
pprint(inspector.get_table_names())
pprint(metadata.tables)


class Log(Base):
    __tablename__ = 'log'
    id = Column(Integer, primary_key=True)
    line_number = Column(Integer, nullable=False)
    time = Column(String, nullable=False)
    date = Column(String, nullable=False)
    comment = Column(String, nullable=False)
    note = Column(String, nullable=True)


class SearchWindow(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(SearchWindow, self).__init__()
        Ui_Form.__init__(self)
        self.setupUi(self)
        # self.show()
        self.log = Log()
        self._load_all_data()
        self.b_refresh.pressed.connect(self._reload)
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)  #EDITS OFF
        self.search_input.textChanged.connect(self._reload)
        self.results_table.setColumnWidth(0, 70)
        self.results_table.setColumnWidth(1, 45)
        self.results_table.setColumnWidth(2, 340)

    def _load_all_data(self):
        search_string = self.search_input.text()
        try:
            # select_this = engine.execute('SELECT * FROM "log" WHERE date="' + this_date + '"')
            if search_string == '':
                select_this = engine.execute('SELECT * FROM "log"')
            else:
                search_string = '%' + search_string + '%'
                select_this = engine.execute('SELECT * FROM "log" WHERE comment LIKE "' + search_string + '"')
            row_number = 0
            last_date = ''
            color_bool = False
            for data in select_this:
                self.results_table.insertRow(row_number)
                self.results_table.setRowHeight(row_number, 2.5)
                self.results_table.setItem(row_number, 0, QtWidgets.QTableWidgetItem())
                self.results_table.setItem(row_number, 1, QtWidgets.QTableWidgetItem())
                self.results_table.setItem(row_number, 2, QtWidgets.QTableWidgetItem())
                date = self.results_table.item(row_number, 0)
                time = self.results_table.item(row_number, 1)
                comment = self.results_table.item(row_number, 2)
                date.setTextAlignment(QtCore.Qt.AlignCenter)
                time.setTextAlignment(QtCore.Qt.AlignCenter)
                date.setText(data.date)
                time.setText(data.time)
                comment.setText(data.comment)
                if last_date != date.text():
                    color_bool = not color_bool
                if color_bool:
                    date.setBackground(QtGui.QColor(208, 236, 249))
                    time.setBackground(QtGui.QColor(208, 236, 249))
                    comment.setBackground(QtGui.QColor(208, 236, 249))
                print(row_number)
                row_number += 1
                last_date = date.text()
        # except exc.InvalidRequestError as e:
        except (exc.InvalidRequestError, Exception) as e:
            print(e)
        # self.results_table.resizeRowsToContents()
        # self.results_table.setWordWrap(True)

    def _reload(self):
        self.results_table.setRowCount(0)
        self._load_all_data()




app = QtWidgets.QApplication([])
application = SearchWindow()
application.setWindowFlags(application.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
application.setWindowIcon(QtGui.QIcon('favicon.ico'))
application.show()
sys.exit(app.exec())