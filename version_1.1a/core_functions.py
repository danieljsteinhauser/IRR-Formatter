import pandas as pd
import numpy as np
import xlrd, openpyxl

#Function to Reformat the Data Inside a Pandas Dataframe and Save as a New Sheet. Original Data is preserved in Another Sheet
def ExcelParse(excelPath, excelSheet, outputPath, hyperionColumns, workOrderColumns):

    #Scan's Document
    df = pd.read_excel(excelPath, excelSheet)

    #Reorders Columns to Match Previous Hyperion Form
    dfHyperion = pd.DataFrame(df, columns = hyperionColumns)
    dfWorkOrder = pd.DataFrame(df, columns = workOrderColumns)

    #Fixes Empty Columns and corrects the "Leaking" header
    dfNaN = dfHyperion.replace(np.nan, '', regex=True)
    dfNaN.rename(columns={'Repair Made to Stop Leak?' : 'Leaking?'}, inplace=True)

    dfWorkOrderNaN = dfWorkOrder.replace(np.nan, '', regex=True)
    dfWorkOrderNaN.rename(columns={'Repair Made to Stop Leak?' : 'Leaking?'}, inplace=True)

    #Exports original data and reformatted data to the original sheet
    with pd.ExcelWriter(outputPath, date_format = 'mm/dd/yy', datetime_format='mm/dd/yy') as writer:
        dfWorkOrderNaN.to_excel(writer, index=False sheet_name=('IRR Basic Information'))
        dfNaN.to_excel(writer, index=False sheet_name=('IRR Reformatted'))
        df.to_excel(writer, index=False, sheet_name=('Original Data'))
        
    del writer
