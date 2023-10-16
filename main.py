import sys
import os
from PyQt5 import QtWidgets, QtGui, QtCore
from main_ui import Ui_MainWindow  # importing our generated file
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QMessageBox, QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal
import xlrd
import xlsxwriter
from openpyxl import Workbook
from send_msg import Send_Msg_thread

class mywindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.lstHistory.sizePolicy().setHorizontalStretch(1)
        self.model = QtGui.QStandardItemModel()
        self.ui.lstHistory.setModel(self.model)
        self.ui.btnStart.clicked.connect(self.on_start)
        self.ui.btnSelectFile.clicked.connect(self.on_selectFile)
        self.ui.btnSelectImgs.clicked.connect(self.on_selectImg)
        self.Send_Msg_thread = None
        self.send_phones = []
        self.send_msgs = []
        self.send_image = None
        self.replace_strs = [" ", "(", ")"]
        self.input_file = None
        self.outPath = None
        self.outFile = []
        self.countryCode = "+33"
        self.msg_cols = []

    def on_selectFile(self):
        filename = QFileDialog.getOpenFileName(self, 'Select Excel file', '/home', "Excel files (*.xls *.xlsx)")
        if(filename == "" or filename == None or filename[0] == "" or filename[0] == None):
            self.ui.edtFile.setText("")
            return
        self.ui.edtFile.setText(filename[0])
        self.input_file = xlrd.open_workbook(filename[0])
        self.ui.cmbSheet.addItems(self.input_file.sheet_names())
        return
    
    def initAllData(self):
        self.model = QtGui.QStandardItemModel()
        self.ui.lstHistory.setModel(self.model)
        self.Send_Msg_thread = None
        self.send_phones = []
        self.send_msgs = []
        self.send_image = None
        self.outPath = None
        self.outFile = []
        self.msg_cols = []
        self.ui.edtMsgCols.setText("")
        self.ui.edtMsgImg.setText("")
        self.ui.btnStart.setEnabled(True)
        return

    def on_selectImg(self):
        filename = QFileDialog.getOpenFileName(self, 'Select Image file', '/home', "Image files (*.png *.jpg)")
        if(filename == "" or filename == None or filename[0] == "" or filename[0] == None):
            self.ui.edtMsgImg.setText("")
            return
        self.ui.edtMsgImg.setText(filename[0])
        self.send_image = filename[0]
        return
        
    def threadStatus(self, type, text):
        if(type != 0):
            item = QtGui.QStandardItem(text)
            self.model.appendRow(item)
            if(type == 2):
                self.exportExcel()
                self.initAllData()
            if type == 3:
                self.initAllData()
        else:
            self.outFile.append(self.xl_sheet.row_values(int(text)))
        return
   
    def exportExcel(self):
        output_file = QFileDialog.getSaveFileName(self,"Save result","result.xlsx","Excel Files(*.xlsx *.xls)")
        if(output_file == "" or output_file == None or output_file[0] == "" or output_file[0] == None):
            return
        workbook = xlsxwriter.Workbook(output_file[0])
        worksheet = workbook.add_worksheet()
        iRow = 0
        iCol = 0
        for rowData in (self.outFile):
            for item in rowData:
                worksheet.write(iRow, iCol,str(item))
                iCol +=1
            iRow += 1
            iCol = 0
        workbook.close()
        return

    def on_start(self):
        self.countryCode = "+"+str(self.ui.spbCountryCode.value())
        if self.input_file is None:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Please select file.')
            msg.setWindowTitle("Error")
            msg.exec()
            return
        sheet_index = self.ui.cmbSheet.currentIndex()
        self.xl_sheet = self.input_file.sheet_by_index(sheet_index)
        self.outFile.append(self.xl_sheet.row_values(0))

        phone_col = self.ui.spbPhone.value()-1
        str_msg_cols = self.ui.edtMsgCols.text()
        if(str_msg_cols == "" or str_msg_cols == None):
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Please input message columns.')
            msg.setWindowTitle("Error")
            msg.exec()
            return
        msg_cols = []
        msg_adds = []
        cols = str_msg_cols.split("]")
        for col in cols:
            strMsg = col.split("[")
            try:
                if len(strMsg) > 1:
                    msg_cols.append(int(strMsg[1])-1)
                    msg_adds.append(strMsg[0])
                else:
                    msg_cols.append(-1)
                    msg_adds.append(strMsg[0])
            except:
                continue


        for i in range(self.xl_sheet.nrows):
            msg_rows = []
            for j in range(len(msg_cols)):
                msgs = msg_adds[j].split("<br>")
                if msg_adds[j].find("<BR>") >= 0:
                    msgs = msg_adds[j].split("<BR>")

                for index in range(len(msgs)):
                    if(index != 0):
                        msg_rows.append("<br>")
                    msg_rows.append(msgs[index])
                
                try:
                    msg_from_xlsx = ""
                    if(int(msg_cols[j]) >= 0):
                        msg_from_xlsx = self.xl_sheet.cell_value(i, int(msg_cols[j]))
                        if msg_from_xlsx.find("http://") >= 0 or msg_from_xlsx.find("https://") >= 0:
                            msg_from_xlsx = " " + msg_from_xlsx + " "
                except:
                    pass
                msg_rows.append(msg_from_xlsx)

            self.send_msgs.append(msg_rows)

        phone_col_data = self.xl_sheet.col_values(phone_col)
        for i in range(len(phone_col_data)):
            if i == 0:
                self.send_phones.append(str(phone_col_data[i]))
                continue

            strPhone =  str(phone_col_data[i])
            if(strPhone == "" or strPhone == None):
                self.send_phones.append("")
                continue
            
            for strr in self.replace_strs:
                strPhone = strPhone.replace(strr, "")
            
            try:
                strPhone = strPhone.split(".")[0]
            except:
                strPhone

            if(strPhone == "" or strPhone == None):
                self.send_phones.append("")
                continue
            
            if(strPhone[0] != "+"):
                if(strPhone[0] == "0"):
                    strPhone = (str(self.countryCode) + strPhone[1:])
                else:
                    strPhone = (str(self.countryCode) + strPhone)
                
            self.send_phones.append(strPhone)

        if(len(self.send_msgs) <= 1):
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Please select correct file.')
            msg.setWindowTitle("Error")
            msg.exec()
            return
        try:
            if self.Send_Msg_thread.driver is not None:
                self.Send_Msg_thread.driver.close()
                self.Send_Msg_thread.driver.quit()
        except:
            print("...")

        self.Send_Msg_thread = Send_Msg_thread()
        self.Send_Msg_thread.output[int, 'QString'].connect(self.threadStatus)
        self.Send_Msg_thread.send_numbers = self.send_phones
        self.Send_Msg_thread.send_msgs = self.send_msgs
        self.Send_Msg_thread.send_image = self.send_image
        self.ui.btnStart.setEnabled(False)
        self.Send_Msg_thread.start()
        return
app = QtWidgets.QApplication([])
application = mywindow()
application.show()
sys.exit(app.exec())