# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'LogHelper\search.ui'
#
# Created by: PyQt5 UI code generator 5.14.1
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(483, 558)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        Form.setWindowOpacity(1.0)
        self.widget = QtWidgets.QWidget(Form)
        self.widget.setGeometry(QtCore.QRect(-1, -1, 483, 559))
        self.widget.setObjectName("widget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.widget)
        self.verticalLayout.setSizeConstraint(QtWidgets.QLayout.SetMaximumSize)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setContentsMargins(10, -1, 10, -1)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.search_input = QtWidgets.QLineEdit(self.widget)
        self.search_input.setMinimumSize(QtCore.QSize(187, 20))
        self.search_input.setObjectName("search_input")
        self.horizontalLayout.addWidget(self.search_input)
        spacerItem = QtWidgets.QSpacerItem(187, 39, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.b_refresh = QtWidgets.QPushButton(self.widget)
        self.b_refresh.setEnabled(True)
        self.b_refresh.setMinimumSize(QtCore.QSize(75, 23))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.b_refresh.setPalette(palette)
        self.b_refresh.setAutoFillBackground(False)
        self.b_refresh.setDefault(False)
        self.b_refresh.setFlat(True)
        self.b_refresh.setObjectName("b_refresh")
        self.horizontalLayout.addWidget(self.b_refresh)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.results_table = QtWidgets.QTableWidget(self.widget)
        self.results_table.setMinimumSize(QtCore.QSize(481, 479))
        self.results_table.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.results_table.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.results_table.setIconSize(QtCore.QSize(5, 5))
        self.results_table.setColumnCount(3)
        self.results_table.setObjectName("results_table")
        self.results_table.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        self.results_table.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        item.setBackground(QtGui.QColor(58, 232, 232, 32))
        self.results_table.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        self.results_table.setHorizontalHeaderItem(2, item)
        self.verticalLayout.addWidget(self.results_table)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setContentsMargins(10, -1, 10, -1)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.results_label = QtWidgets.QLabel(self.widget)
        self.results_label.setMinimumSize(QtCore.QSize(461, 23))
        self.results_label.setObjectName("results_label")
        self.horizontalLayout_2.addWidget(self.results_label)
        self.verticalLayout.addLayout(self.horizontalLayout_2)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Search"))
        self.search_input.setPlaceholderText(_translate("Form", "Filter"))
        self.b_refresh.setText(_translate("Form", "Refresh"))
        item = self.results_table.horizontalHeaderItem(0)
        item.setText(_translate("Form", "Date"))
        item = self.results_table.horizontalHeaderItem(1)
        item.setText(_translate("Form", "Time"))
        item = self.results_table.horizontalHeaderItem(2)
        item.setText(_translate("Form", "Comment"))
        self.results_label.setText(_translate("Form", "results"))