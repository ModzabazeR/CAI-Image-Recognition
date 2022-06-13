# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Users\Kittiphat\Documents\CAI_C\GUI\2.0.ui'
#
# Created by: PyQt5 UI code generator 5.15.6
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


import datetime
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QFileDialog
import sys, os, time, re, utils, time
import threading
import InvoiceExtract as Ie

output = ''

class Ui_Main(object):
    
    
    def setupUi(self, Main):
        
        Main.setObjectName("Main")
        Main.setWindowModality(QtCore.Qt.ApplicationModal)
        Main.setWindowIcon(QtGui.QIcon('icon.ico'))
        Main.resize(640, 480)
        Main.setMinimumSize(QtCore.QSize(640, 480))
        Main.setMaximumSize(QtCore.QSize(640, 480))
        
        self.output = ''
        
        self.exportButton = QtWidgets.QPushButton(Main)
        self.exportButton.setEnabled(False)
        self.exportButton.setGeometry(QtCore.QRect(530, 430, 93, 28))
        self.exportButton.setObjectName("exportButton")
        self.exportButton.clicked.connect(self.start_submit_thread)
        
        self.io = QtWidgets.QGroupBox(Main)
        self.io.setEnabled(True)
        self.io.setGeometry(QtCore.QRect(10, 20, 621, 141))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(129)
        sizePolicy.setVerticalStretch(66)
        sizePolicy.setHeightForWidth(self.io.sizePolicy().hasHeightForWidth())
        self.io.setSizePolicy(sizePolicy)
        self.io.setObjectName("io")
        
        self.horizontalLayoutWidget = QtWidgets.QWidget(self.io)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(120, 30, 491, 41))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        
        self.inputEdit = QtWidgets.QLineEdit(self.horizontalLayoutWidget)
        self.inputEdit.setObjectName("inputEdit")
        self.horizontalLayout.addWidget(self.inputEdit)
        self.inputButton = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.inputButton.setObjectName("inputButton")
        self.inputFileLists =self.inputButton.clicked.connect(self.getFileNamesInput)
        self.horizontalLayout.addWidget(self.inputButton)
        
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.io)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(10, 30, 111, 81))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label.setObjectName("label")
        self.verticalLayout_2.addWidget(self.label)
        self.label_2 = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_2.addWidget(self.label_2)
        self.horizontalLayoutWidget_2 = QtWidgets.QWidget(self.io)
        self.horizontalLayoutWidget_2.setGeometry(QtCore.QRect(120, 70, 491, 41))
        self.horizontalLayoutWidget_2.setObjectName("horizontalLayoutWidget_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.outputEdit = QtWidgets.QLineEdit(self.horizontalLayoutWidget_2)
        self.outputEdit.setEnabled(True)
        self.outputEdit.setObjectName("outputEdit")
        self.horizontalLayout_2.addWidget(self.outputEdit)
        self.outputButtion = QtWidgets.QPushButton(self.horizontalLayoutWidget_2)
        self.outputButtion.setObjectName("outputButtion")
        self.outputFolder = self.outputButtion.clicked.connect(self.getFileNamesOutput)
        self.horizontalLayout_2.addWidget(self.outputButtion)
        self.groupBox = QtWidgets.QGroupBox(Main)
        self.groupBox.setGeometry(QtCore.QRect(10, 180, 621, 239))
        self.groupBox.setObjectName("groupBox")
        self.statusScrollArea = QtWidgets.QScrollArea(self.groupBox)
        self.statusScrollArea.setGeometry(QtCore.QRect(9, 26, 601, 201))
        self.statusScrollArea.setWidgetResizable(True)
        self.statusScrollArea.setObjectName("statusScrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 599, 199))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.statusScrollArea.setWidget(self.scrollAreaWidgetContents)
        
        self.textBrowser = QtWidgets.QTextBrowser(self.scrollAreaWidgetContents)
        self.textBrowser.setGeometry(QtCore.QRect(0, 0, 601, 201))
        self.textBrowser.setObjectName("textLog")
        self.statusScrollArea.setWidget(self.scrollAreaWidgetContents)
        
    
        self.retranslateUi(Main)
        QtCore.QMetaObject.connectSlotsByName(Main)

    def retranslateUi(self, Main):
        _translate = QtCore.QCoreApplication.translate
        Main.setWindowTitle(_translate("Main", "DocJuice! 2.0"))
        self.exportButton.setText(_translate("Main", "Export"))
        self.io.setTitle(_translate("Main", "Settings"))
        self.inputButton.setText(_translate("Main", "Browse..."))
        self.label.setText(_translate("Main", "PDF files to export: "))
        self.label_2.setText(_translate("Main", "Output directory: "))
        self.outputButtion.setText(_translate("Main", "Browse..."))
        self.groupBox.setTitle(_translate("Main", "Status"))
        self.textBrowser.setText(_translate("Main", r"%s : Welcome to DocJuice 2.0! "%datetime.datetime.now().strftime("%H:%M:%S")))
        self.inputEdit.setReadOnly(True)
        self.inputEdit.setStyleSheet("color: #252525; background-color: #F0F0F0;")
        self.outputEdit.setReadOnly(True)
        self.outputEdit.setStyleSheet("color: #252525; background-color: #F0F0F0;")
        
    def getFileNamesInput(self):
        try:
            file_filter = 'PDF File (*.pdf)'
            response = QFileDialog.getOpenFileNames(
                caption='Select PDF file to export',
                directory=os.getcwd(),
                filter=file_filter,
                initialFilter='PDF File (*.pdf)')
            print(response[0])
            self.file_list = response[0]
            if response[0] :
                self.inputEdit.setText(f"Selected {len(response[0])} PDF files to process.")
                self.textBrowser.append(f"{datetime.datetime.now().strftime('%H:%M:%S')}: You have selected {len(response[0])} files. - {', '.join(map(os.path.basename, response[0]))}")
                return response[0]
        except UnicodeError as e:
            self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s "%(datetime.datetime.now().strftime("%H:%M:%S"),e)))
        
    def getFileNamesOutput(self):
        response = QFileDialog.getExistingDirectory(
            caption='Select output directory',
            directory=os.getcwd(),
            )
        print(response)
        if response[0] :
            self.outputEdit.setText(response)
            self.dataCheck()
            self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : Set output directory at %s "%(datetime.datetime.now().strftime("%H:%M:%S"),response)))

    def dataCheck(self):
        if self.inputEdit.text() != '' and self.outputEdit.text() != '':
            self.exportButton.setEnabled(True)
            print('output is : ', self.outputEdit.text())
    
    def getTime(self):
        named_tuple = time.localtime() # get struct_time
        time_string = time.strftime("%H:%M:%S", named_tuple)
        return time_string
    
    def whenClick(self):
        global output
        self.exportButton.setEnabled(False)
        self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : Exporting... "%(self.getTime())))
        self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r""))
        
        ### Start export ###
        # export script here 
        # input file destination at 'self.inputEdit.text()'
        # output file destination at 'self.outputEdit.text()'
        
        
        input = self.file_list
        Ie.output = self.outputEdit.text()
        
        len_file = len(input)
        self.file_scanned = 0
        self.skipped = 0
        for file in input :
            cur_file = os.path.basename(file)
            print(f"Processing {cur_file} ...")
            self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : Processing %s [%d/%d]"%(self.getTime(), cur_file, input.index(file)+1, len_file)))
            if file.lower().endswith(".pdf"):
                if all(x in utils.ie_extract_text(file) for x in Ie.KBANK_KEYWORDS):
                    inv = Ie.KBANKInvoice(file)
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s detected as KBANK invoice "%(self.getTime(), cur_file)))
                    print (f"{cur_file} detected as KBANK invoice")
                elif all(x in utils.ie_extract_text(file) for x in Ie.BBL_KEYWORDS):
                    inv = Ie.BBLInvoice(file)
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s detected as BBL invoice "%(self.getTime(), cur_file)))
                    print (f"{cur_file} detected as BBL invoice")
                elif all(x in utils.ie_extract_text(file) for x in Ie.SCB_KEYWORDS):
                    inv = Ie.SCBInvoice(file)
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s detected as SCB invoice "%(self.getTime(), cur_file)))
                    print (f"{cur_file} detected as SCB invoice")
                elif all(x in utils.ie_extract_text(file) for x in Ie.TTB_KEYWORDS) and re.search(Ie.TTB_PATTERN, utils.ie_extract_text(file)):
                    inv = Ie.TTBInvoice(file)
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s detected as TTB invoice "%(self.getTime(), cur_file)))
                    print (f"{cur_file} detected as TTB invoice")
                elif all(x in utils.ie_extract_text(file) for x in Ie.BAY_KEYWORDS):
                    inv = Ie.BAYInvoice(file)
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s detected as BAY invoice "%(self.getTime(), cur_file)))
                    print (f"{cur_file} detected as BAY invoice")
                else:
                    self.skipped += 1
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s not supported "%(self.getTime(), cur_file)))
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r""))
                    continue 
                try:
                    # inv.get_invoice_info()
                    inv.to_excel()
                    self.file_scanned += 1
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s Completed "%(self.getTime(), cur_file)))
                except Exception as e:
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", f"Error at {cur_file}: {e}"))
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : Skipped file %s "%(self.getTime(), cur_file)))
                    self.skipped += 1
                    continue
                finally:
                    inv.close()
                    self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r""))
                    self.textBrowser.moveCursor(QtGui.QTextCursor.End)
        
        self.scriptFinish()

                    
                    
        
    def scriptFinish(self):
        self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %d files exported, %d files skipped "%(self.getTime(), self.file_scanned, self.skipped)))
        self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : Finished export "%(self.getTime())))

        if self.skipped != len(self.file_list):
            self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : Merging Excel Files... "%(self.getTime())))
            status = Ie.compile_workbooks(self.outputEdit.text(), f"final_{time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
            self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : %s "%(self.getTime(), status)))

        with open("log.txt", "w", encoding="utf-8") as f:
            f.write(self.textBrowser.toPlainText())
        self.textBrowser.append(QtCore.QCoreApplication.translate("Main", r"%s : Log file saved to %s "%(self.getTime(), "log.txt")))
        self.textBrowser.moveCursor(QtGui.QTextCursor.End)

        time.sleep(2)
        self.exportButton.setEnabled(True)
        
    def start_submit_thread(self):
        self.file_scanned = 0
        self.skipped = 0
        process_thread = threading.Thread(target=self.whenClick)
        process_thread.daemon = True
        process_thread.start()



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Main = QtWidgets.QWidget()
    ui = Ui_Main()
    ui.setupUi(Main)
    Main.show()
    sys.exit(app.exec_())
