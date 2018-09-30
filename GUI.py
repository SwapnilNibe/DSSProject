from PyQt5 import QtCore, QtGui, QtWidgets
import os
import openpyxl as xl

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(690, 552)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(270, 210, 111, 31))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(270, 260, 111, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(190, 150, 401, 21))
        self.lineEdit.setObjectName("lineEdit")
        self.btn = QtWidgets.QPushButton(self.centralwidget)
        self.btn.setGeometry(QtCore.QRect(60, 150, 75, 23))
        self.btn.setObjectName("pushButton")
        self.btn.clicked.connect(self.showDialog)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 690, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.pushButton_2.clicked.connect(self.launchfile)
        self.pushButton.clicked.connect(self.createfile)
    	
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Create Excel File"))
        self.pushButton_2.setText(_translate("MainWindow", "Open Excel "))
        self.lineEdit.setText(_translate("MainWindow", "There is no current directory"))
        self.btn.setText(_translate("MainWindow", "Edit Directory"))
        
    def showDialog(self):
        text,result = QtWidgets.QInputDialog.getText(None ,'Working Directory','Enter path')
        if result == True:
            self.lineEdit.setText(str(text))
        os.chdir(text.replace("\\","\\\\"))
            
    def createfile(self):
        file = xl.Workbook()
        file.save("SalesDSS.xlsx")        
        file.close()

    def launchfile(self):
        os.startfile("SalesDSS.xlsx")    	


if __name__ == "__main__":
	import sys
	app = QtWidgets.QApplication(sys.argv)
	MainWindow = QtWidgets.QMainWindow()
	ui = Ui_MainWindow()
	ui.setupUi(MainWindow)
	MainWindow.show()
	sys.exit(app.exec_())


