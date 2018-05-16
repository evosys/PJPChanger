#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os, sys, io, time, signal, inspect
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui import Ui_PJPChanger
from lxml import etree
import pyodbc
import appinfo
from pathlib import Path

NEWDIR           = 'XML-output'
EXIT_CODE_REBOOT = -23467876230

_db    = "Centegy_SnDPro_UID"
_uname = "sa"
_pwd   = "unilever1"
_configFile = "PJP_Changer.ini"

# main class
class mainWindow(QMainWindow, Ui_PJPChanger) :
    def __init__(self) :
        QMainWindow.__init__(self)
        self.setupUi(self)

        # app icon
        self.setWindowIcon(QIcon(':/resources/icon.png'))

        # centering app
        tr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        tr.moveCenter(cp)
        self.move(tr.topLeft())

        # Set database result as global variabel
        self.rowDataHCF = self.getDataSales(103)
        self.rowDataPC = self.getDataSales(102)

        # populate combo box HCF
        for rSalesHCF in self.rowDataHCF:
            self.cbHCF.addItem(rSalesHCF[0], rSalesHCF[2])

        for rSalesPC in self.rowDataPC:
            self.cbPC.addItem(rSalesPC[0], rSalesPC[2])

        # other components
        self.lbPath.hide()
        self.lbPath.clear()
        self.btOpen.clicked.connect(self.openXml)
        self.btSave.clicked.connect(self.saveChange)
        self.edFile.textChanged.connect(self.setItem)
        # self.toolButton.clicked.connect(self.FrmAddrs)

        # checking chekbox
        self.ckHCF.stateChanged.connect(self.chkHCF)
        self.ckPC.stateChanged.connect(self.chkPC)
        # self.ckDC.stateChanged.connect(self.chkDC)
        # self.ckXC.stateChanged.connect(self.chkXC)


    # Frominput address server
    def FrmAddrs(self):

        qip = QtWidgets.QInputDialog()

        # define default setting value
        settings = QSettings(_configFile, "PJP_Changer")
        DefSrvr = settings.value('server')

        text, pressOK = qip.getText(self, "PJP Changer","Please enter the hostname of the database server:", QLineEdit.Normal, settings.value('server'))
        if pressOK :
            if text != "" :
                # save setting value
                settings.setValue("server", text)
                self.restart()
            else :
                reply = QMessageBox.critical(self, "Error", "Please input address database server.", QMessageBox.Ok)
                self.FrmAddrs()


    # checkBox HCF
    def chkHCF(self, state) :
        if state == QtCore.Qt.Checked :
            self.cbHCF.setEnabled(True)
            return True
        else :
            self.cbHCF.setEnabled(False)
            return False


    # checkBox PC
    def chkPC(self, state) :
        if state == QtCore.Qt.Checked :
            self.cbPC.setEnabled(True)
            return True
        else :
            self.cbPC.setEnabled(False)
            return False

    # checkBox DC
    # def chkDC(self, state) :
        # if state == QtCore.Qt.Checked :
            # self.cbDC.setEnabled(True)
            # return True
        # else :
            # self.cbDC.setEnabled(False)
            # return False

    # checkBox XC
    # def chkXC(self, state) :
        # if state == QtCore.Qt.Checked :
            # self.cbXC.setEnabled(True)
            # return True
        # else :
            # self.cbXC.setEnabled(False)
            # return False


    # search data HCF in array
    def SearchDataHCF(self, val) :
        for findDat in self.rowDataHCF :
            for data in findDat :
                if data == val :
                    tmp = list(findDat)
                    return tmp


    # search data PC in array
    def SearchDataPC(self, val) :
        for findDat in self.rowDataPC :
            for data in findDat :
                if data == val :
                    tmp = list(findDat)
                    return tmp


    # connect to DB
    def connDB(self) :
        settings = QSettings(_configFile, "PJP_Changer")
        server = settings.value("server")

        try:
            cnxn = pyodbc.connect(driver='{ODBC Driver 13 for SQL Server}',
                                  server=server,
                                  database=_db,
                                  uid=_uname,
                                  pwd=_pwd,
                                  timeout=5)

            cnxn.setencoding(encoding='utf-8', ctype=pyodbc.SQL_CHAR)

            return cnxn

        except pyodbc.Error as err :
            msg = "Can't connect to database server.<br>Do you want to input address server manually?"
            errorSrv = QMessageBox.critical(self, "Error", msg, QMessageBox.Yes | QMessageBox.Abort)

            if errorSrv == QMessageBox.Yes :
                self.FrmAddrs()
                sys.exit(0)
            else :
                raise SystemExit(0)


    # Get data items by category
    def getDataItems(self, CatGor) :

        cursor = None

        cnxn = self.connDB()
        cursor = cnxn.cursor()

        que = "SELECT a.SKU FROM SKU_CATEGORY a INNER JOIN SELLING_CATEGORY b on b.SELL_CATEGORY = a.SELL_CATEGORY WHERE a.SELL_CATEGORY = ?"

        params = str(CatGor)

        cursor.execute(que, params)

        results = cursor.fetchall()

        cursor.close()
        del cursor
        cnxn.close()

        items = []

        for element in results :
            for item in element :
                items.append(item)

        return items


    # get data from database HCF
    def getDataSales(self, CatGor):

        settings = QSettings(_configFile, "PJP_Changer")
        server = settings.value("server")

        cursor = None

        cnxn = self.connDB()
        cursor = cnxn.cursor()

        que = "SELECT CONCAT(p.PJP,' / '+P.LDESC) as PJPSales, p.PJP, p.DSR, p.LDESC as SalesName, a.LDESC as CategoryName FROM PJP_HEAD p INNER JOIN SELLING_CATEGORY a on a.SELL_CATEGORY = p.SELL_CATEGORY INNER JOIN DSR ds on ds.DSR = p.DSR WHERE a.SELL_CATEGORY = ? and ds.JOB_TYPE = 01 and p.ACTIVE = 1"

        params = str(CatGor)

        cursor.execute(que, params)

        results = cursor.fetchall()

        cursor.close()
        del cursor
        cnxn.close()

        self.statusBar().showMessage('Connected: '+settings.value("server", type=str)+ ' (v'+appinfo._version+')')

        return results


    # open XML
    def openXml(self):
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","XML Files (*.xml)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")

            # enable checkbox
            self.ckHCF.setEnabled(True)
            self.ckPC.setEnabled(True)
            # self.ckDC.setEnabled(True)
            # self.ckXC.setEnabled(True)

        # End of def openXML


    # set
    def setItem(self):
        path = self.lbPath.text()

        if len(path) == 0:
            # set index combobox
            self.cbHCF.setCurrentIndex(0)
            self.cbPC.setCurrentIndex(0)
            # self.cbDC.setCurrentIndex(0)
            # self.cbXC.setCurrentIndex(0)

        else:
            tree = etree.parse(path)
            vPJP = tree.find('.//RouteCode').text
            vSales = tree.find('.//SalesmanCode').text

        # get value of combobox, set it with same value of XML file
        indexSales = self.cbHCF.findData(vSales)
        if indexSales >= 0:

            self.cbHCF.setCurrentIndex(indexSales)
            self.cbPC.setCurrentIndex(indexSales)
            # self.cbDC.setCurrentIndex(indexSales)
            # self.cbXC.setCurrentIndex(indexSales)

        # End of def setItem.


    # changer
    def changer(self, pathXML, cmpCat) :

        if len(pathXML) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XML file first!", QMessageBox.Ok)
        else :

            # for HCF
            if cmpCat == '103' :
                codeSales = str(self.cbHCF.itemData(self.cbHCF.currentIndex()))
                nameSales = str(self.cbHCF.currentText())
                # search sales PJP
                srchTMP = self.SearchDataHCF(nameSales)
                countDoc = "3"

            # for PC
            if cmpCat == '102' :
                codeSales = str(self.cbPC.itemData(self.cbPC.currentIndex()))
                nameSales = str(self.cbPC.currentText())
                # search sales PJP
                srchTMP = self.SearchDataPC(nameSales)
                countDoc = "2"

            # for DC
            # if cmpCat == '3' :
                # codeSales = str(self.cbDC.itemData(self.cbDC.currentIndex()))
                # nameSales = str(self.cbDC.currentText())

            # for XC
            # if cmpCat == '4' :
                # codeSales = str(self.cbXC.itemData(self.cbXC.currentIndex()))
                # nameSales = str(self.cbXC.currentText())

            # Parsing file xml
            tree = etree.parse(pathXML)

            # Validating the value before make any change
            if len(codeSales) > 0 :
                suffix = "ORD"

                # get search
                srch = srchTMP

                # format 4 Digit
                codePJP = "{:0>4}".format(srch[1])

                # get document number
                DocNum = tree.find('.//DocumentNumber').text
                DocNum = DocNum[:2] + countDoc + DocNum[3:] # replace 3 digit from front

                # get RVTKey
                RVTKey = tree.find('.//RVTKey').text
                RVTKey = RVTKey[:2] + countDoc + RVTKey[3:] # replace 3 digit from front

                # change RouteCode
                tree.find('.//RouteCode').text = codePJP

                # change BillToCustomer
                tree.find('.//BillToCustomer').text = codePJP

                # change Document number
                for node in tree.xpath('.//DocumentNumber') :
                    node.text = DocNum

                # change RVTKey
                for node in tree.xpath('.//RVTKey') :
                    node.text = RVTKey

                # change DocumentPrefix
                for node in tree.xpath('.//DocumentPrefix') :
                    node.text = codePJP + suffix

                tree.find('.//SalesmanCode').text = codeSales

                # get items by category
                arrItems = self.getDataItems(cmpCat)

                for order in tree.xpath("//SalesOrderDetail") :
                    item = order.xpath('ItemCode')
                    item_code = item[0].text

                    if item_code not in arrItems:
                        order.getparent().remove(order)

            else:
                QMessageBox.warning(self, "Warning", "Please select Salesman first", QMessageBox.Ok)

            # write data XML with Value of combobox.
            if len(codePJP) > 0 and len(codeSales) > 0:
                resPath, resFilename = os.path.split(pathXML)
                current_dir = os.getcwd()
                Fname, Fext = os.path.splitext(resFilename)
                newFile = "{Fname}{cmpCat}{Fext}".format(Fname=Fname, cmpCat=cmpCat, Fext=Fext)
                resPathFile = os.path.abspath(os.path.join(current_dir, NEWDIR, newFile))
                resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR, newFile)))
                resultPath.parent.mkdir(parents=True, exist_ok=True)

                tree.write(resPathFile, xml_declaration=True, encoding='utf-8', method="xml")
                return True
            else :
                return False
            # End of def saveChange.


    # button save
    def saveChange(self) :
        # get path directory
        pathXML = self.lbPath.text()
        # set path to variabel
        resPath, resFilename = os.path.split(pathXML)
        current_dir = os.getcwd()
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        if len(pathXML) == 0:

            QMessageBox.warning(self, "Warning", "Please select XML file first!", QMessageBox.Ok)

        else :

            if self.ckHCF.isChecked() or self.ckPC.isChecked() or self.ckDC.isChecked() or self.ckXC.isChecked():

                if self.ckHCF.isChecked() :
                    catPO = '103'
                    self.changer(pathXML, catPO)

                if self.ckPC.isChecked() :
                    catPO = '102'
                    self.changer(pathXML, catPO)

                # if self.ckDC.isChecked() :
                    # catPO = '3'
                    # self.changer(pathXML, catPO)

                # if self.ckXC.isChecked() :
                    # catPO = '4'
                    # self.changer(pathXML, catPO)

                reply = QMessageBox.information(self, "Information", "Success!", QMessageBox.Ok)
                if reply == QMessageBox.Ok :
                    os.startfile(str(resultPath))

            else :

                QMessageBox.warning(self, "Warning", "Please choose one or more options.", QMessageBox.Ok)


    # restart apps
    def restart(self) :
        QApplication.exit(EXIT_CODE_REBOOT)
        QProcess.startDetached(QCoreApplication.applicationFilePath())


if __name__ == '__main__' :
    app = QApplication(sys.argv)

    # create splash screen
    splash_pix = QPixmap(':/resources/unilever_splash.png')

    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setWindowFlags(QtCore.Qt.FramelessWindowHint)
    splash.setEnabled(False)

    # adding progress bar
    progressBar = QProgressBar(splash)
    progressBar.setMaximum(10)
    progressBar.setGeometry(17, splash_pix.height() - 20, splash_pix.width(), 50)

    splash.show()

    for iSplash in range(1, 11) :
        progressBar.setValue(iSplash)
        t = time.time()
        while time.time() < t + 0.1 :
            app.processEvents()

    time.sleep(1)

    window = mainWindow()
    window.setWindowTitle(appinfo._appname)
    # window.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
    # window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)
    window.show()
    splash.finish(window)
    sys.exit(app.exec_())