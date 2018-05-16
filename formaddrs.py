# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'formaddrs.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_FrmAddrs(object):
    def setupUi(self, FrmAddrs):
        FrmAddrs.setObjectName("FrmAddrs")
        FrmAddrs.resize(280, 111)
        self.centralwidget = QtWidgets.QWidget(FrmAddrs)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(60, 55, 161, 27))
        self.pushButton.setObjectName("pushButton")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(20, 20, 241, 22))
        self.lineEdit.setObjectName("lineEdit")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 0, 231, 20))
        self.label.setObjectName("label")
        FrmAddrs.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(FrmAddrs)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 280, 21))
        self.menubar.setObjectName("menubar")
        FrmAddrs.setMenuBar(self.menubar)

        self.retranslateUi(FrmAddrs)
        QtCore.QMetaObject.connectSlotsByName(FrmAddrs)

    def retranslateUi(self, FrmAddrs):
        _translate = QtCore.QCoreApplication.translate
        FrmAddrs.setWindowTitle(_translate("FrmAddrs", "MainWindow"))
        self.pushButton.setText(_translate("FrmAddrs", "Save"))
        self.label.setText(_translate("FrmAddrs", "Please enter address for database server:"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    FrmAddrs = QtWidgets.QMainWindow()
    ui = Ui_FrmAddrs()
    ui.setupUi(FrmAddrs)
    FrmAddrs.show()
    sys.exit(app.exec_())

