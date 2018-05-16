import os, sys, io, time
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from pjpchanger import Ui_PJPChanger
from lxml import etree
import pyodbc
from pathlib import Path

class mainWindow(QMainWindow, Ui_PJPChanger):
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
        self.rowData = self.getData()

        # populate combo box HCF
        for rSales in self.rowData:
            self.cbHCF.addItem(rSales[2], rSales[1])
            self.cbPC.addItem(rSales[2], rSales[1])
            # self.cbDC.addItem(rSales[2], rSales[1])
            # self.cbXC.addItem(rSales[2], rSales[1])

        # other components
        self.lbPath.hide()
        self.btOpen.clicked.connect(self.openXml)
        self.btSave.clicked.connect(self.saveChange)
        self.edFile.textChanged.connect(self.setItem)

        # checking chekbox
        self.ckHCF.stateChanged.connect(self.chkHCF)
        self.ckPC.stateChanged.connect(self.chkPC)
        # self.ckDC.stateChanged.connect(self.chkDC)
        # self.ckXC.stateChanged.connect(self.chkXC)


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

    # search data in array
    def SearchData(self, val) :
        for findDat in self.rowData :
            for data in findDat :
                if data == val :
                    tmp = list(findDat)
                    return tmp

    # get data from database
    def getData(self):
        # Getting Data Section
        # server = 'localhost\SQLEXPRESS'
        # database = 'Centegy_SnDPro_UID'
        # username = 'sa'
        # password = 'unilever1'

        server = 'EVOSYS137\SQLEXPRESS'
        database = 'Centegy_SnDPro_UID'
        username = 'sa'
        password = 'unilever1'

        # server = 'den1.mssql4.gear.host'
        # database = 'sqlsrv'
        # username = 'sqlsrv'
        # password = 'terserah!'

        try:
            cnxn = pyodbc.connect(driver='{ODBC Driver 13 for SQL Server}',
                                  server=server,
                                  database=database,
                                  uid=username,
                                  pwd=password,
                                  timeout=5)

            cnxn.setencoding(encoding='utf-8', ctype=pyodbc.SQL_CHAR)
            cursor = cnxn.cursor()
        except pyodbc.Error as err :
            # self.hide()
            QMessageBox.critical(self, "Error", "Can't connect to database server.", QMessageBox.Abort)

            raise SystemExit(0)

        cursor.execute("SELECT PJP, DSR, LDESC, SELL_CATEGORY FROM PJP_HEAD WHERE ACTIVE = 1")

        results = cursor.fetchall()

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

    # change 2 digit from first c

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

    # logo path
    def resource_path(self, relative_path) :
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    def changer(self, pathXML, cmpCat) :

        if len(pathXML) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XML file first!", QMessageBox.Ok)
        else :

            # for HCF
            if cmpCat == '1' :
                codeSales = str(self.cbHCF.itemData(self.cbHCF.currentIndex()))
                nameSales = str(self.cbHCF.currentText())

            # for PC
            if cmpCat == '2' :
                codeSales = str(self.cbPC.itemData(self.cbPC.currentIndex()))
                nameSales = str(self.cbPC.currentText())

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

                # search data
                srch = self.SearchData(nameSales)

                # format 4 Digit
                codePJP = "{:0>4}".format(srch[0])

                # get document number
                DocNum = tree.find('.//DocumentNumber').text
                DocNum = DocNum[:2] + cmpCat + DocNum[3:] # replace 3 digit from front

                # get RVTKey
                RVTKey = tree.find('.//RVTKey').text
                RVTKey = RVTKey[:2] + cmpCat + RVTKey[3:] # replace 3 digit from front

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

            else:
                QMessageBox.warning(self, "Warning", "Please select Salesman first", QMessageBox.Ok)

            # write data XML with Value of combobox.
            if len(codePJP) > 0 and len(codeSales) > 0:
                resPath, resFilename = os.path.split(pathXML)
                current_dir = os.getcwd()
                Fname, Fext = os.path.splitext(resFilename)
                newDir = 'output_XML'
                newFile = "{Fname}{cmpCat}{Fext}".format(Fname=Fname, cmpCat=cmpCat, Fext=Fext)
                resPathFile = os.path.abspath(os.path.join(current_dir, newDir, newFile))
                resultPath = Path(os.path.abspath(os.path.join(current_dir, newDir, newFile)))
                resultPath.parent.mkdir(parents=True, exist_ok=True)

                tree.write(resPathFile, xml_declaration=True, encoding='utf-8', method="xml")
                return True
            else :
                return False
            # End of def saveChange.


    # button save
    def saveChange(self) :

        pathXML = self.lbPath.text()
        # get path directory
        resPath, resFilename = os.path.split(pathXML)
        current_dir = os.getcwd()
        newDir = 'output_XML'
        resultPath = Path(os.path.abspath(os.path.join(current_dir, newDir)))

        if len(pathXML) == 0:

            QMessageBox.warning(self, "Warning", "You Must Select File First!", QMessageBox.Ok)

        else :

            if self.ckHCF.isChecked() or self.ckPC.isChecked() or self.ckDC.isChecked() or self.ckXC.isChecked():

                if self.ckHCF.isChecked() :
                    catPO = '1'
                    self.changer(pathXML, catPO)

                if self.ckPC.isChecked() :
                    catPO = '2'
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
    # window.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
    # window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)
    window.show()
    splash.finish(window)

    sys.exit(app.exec_())