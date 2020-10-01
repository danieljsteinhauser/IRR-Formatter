import pandas as pd
import numpy as np
import xlrd, openpyxl
import sys
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QWidget, QApplication, QFileDialog, QMessageBox
import gui
import core_functions as f

# =======================================================================
# List Templates for Excel Sheets
# =======================================================================
hyperionColumns =['EC','Ar Subject','Remarks','Common Trench','EFV Valve', 'Main Depth','Main Loc',"Main Mat",'Main Pressure',
'Main Size','MB Length', 'MB Insert','Repair Date', 'Repair Made to Stop Leak?', #changed to leaking later
'Direct Bury',"MB Mat",'MB Old Size','MB Size','Work Order Nbr','Ar Number','Crew Leader','BR Enters',
'Pipe Test Pressure','Outside Riser','BR To','Svc Supply','Tap Loc1', 'Tap Loc2','Tap Size', 'Valve Loc1', 'Valve Loc2']

workOrderColumns =['EC', 'Ar Subject', 'Work Order Nbr', 'Ar Number', 'Repair Date', 'Crew Leader', 'Repair Made to Stop Leak?', 'Remarks', 'New Remarks', 'Svc Supply', 'Repair Method']  

# =======================================================================
# create app and widget window + dialog GUI 
# =======================================================================
app = QApplication(sys.argv)
window = QWidget()

ui = gui.Ui_formMain() #instantiation of my Ui_Form class
ui.setupUi(window)

# =======================================================================
# GUI Event Handlers and Objects 
# =======================================================================

def submitInput():
    global outputPath
    global excelSheet

    try:
        output, _ = QFileDialog.getSaveFileName(window, "Save file Location")
        outputPath = output + ".xlsx"
        excelSheet = ui.comboBox.currentText()
        f.ExcelParse(excelPath, excelSheet, outputPath, hyperionColumns, workOrderColumns)

        #Completion Notice
        QMessageBox.information(window, "Success!", "The Excel file has been sucessfully reformatted and placed in\n"+ outputPath)
    except:
        QMessageBox.information(window, "An Unexpected Error Has Occured!", "Error: 0002\n\nFailure encountered during file save. Please ensure no special characters were used during the naming of this file and try again. If error persists, please contact Daniel Steinhauser.") 

def DefineInput():
    
    global excelPath

    try:
        excelPath, _ = QFileDialog.getOpenFileName(window, "Select an Excel File","","Excel (*.xlsx)")
        ui.comboBox.clear()
        excelSheetsUnformatted = pd.ExcelFile(excelPath).sheet_names
        excelSheets = sorted(excelSheetsUnformatted)
        ui.comboBox.addItems(excelSheets)
    except:
        pass


# =======================================================================
# connect signals 
# =======================================================================

#Pushbutton Submit
ui.pbAccept.clicked.connect(submitInput)


#Pushbutton Select New File
ui.pbNewFile.clicked.connect(DefineInput)

# =======================================================================
# run app 
# =======================================================================
window.show()
sys.exit(app.exec_())

