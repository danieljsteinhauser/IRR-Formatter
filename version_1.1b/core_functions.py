import pandas as pd
import numpy as np
import xlrd, openpyxl

#Function to Reformat the Data Inside a Pandas Dataframe and Save as a New Sheet. Original Data is preserved in Another Sheet
def ExcelParse(excelPath, excelSheet, hyperionColumns):

    #Scan's Document
    df = pd.read_excel(excelPath, excelSheet)

    #Reorders Columns to Match Previous Hyperion Form
    dfHyperion = pd.DataFrame(df, columns = hyperionColumns)

    #Fixes Empty Columns and corrects the "Leaking" header
    dfNaN = dfHyperion.replace(np.nan, '', regex=True)
    dfNaN.rename(columns={'Repair Made to Stop Leak?' : 'Leaking?'}, inplace=True)

    #Exports original data and reformatted data to the original sheet
    writer = pd.ExcelWriter(excelPath, date_format = 'mm/dd/yy', datetime_format='mm/dd/yy', engine='openpyxl')
    writer.book = openpyxl.load_workbook(excelPath)
    del writer.book[excelSheet]
    dfNaN.to_excel(writer, index=False, sheet_name=excelSheet)
    writer.save()
        
    del writer
