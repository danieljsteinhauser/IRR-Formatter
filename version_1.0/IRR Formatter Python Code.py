import pandas as pd
import numpy as np
import xlrd, openpyxl
from tkinter import filedialog, messagebox, Tk 


#Reformatted Order for Excel Document 

columnList =[
'EC',
'Ar Subject',
'Remarks',
'Common Trench',
'EFV Valve', 
'Main Depth', 
'Main Loc',
"Main Mat'l",
'Main Pressure',
'Main Size',
'MB Length', 
'MB Insert',
'Repaired Date Formatted',
'Repair Made?', #changed to leaking later
'Direct Bury',
"MB Mat'l",
'MB Old Size',
'MB Size',
'Work Order Nbr',
'Ar Number',    
'Crew Leader',  
'BR Enters',
'Pipe Test Pressure',
'Outside Riser',    
'BR To',
'Svc Supply',
'Tap Loc 1', 'Tap Loc 2',
'Tap Size',
'Valve Loc 1', 'Valve Loc 2',

]

#Function to Reformat the Data Inside a Pandas Dataframe and Save as a New Sheet. Original Data is preserved in Another Sheet
def ExcelParse(columnList):
    filePath = filedialog.askopenfilename(title = "Select file", filetypes = (("Excel Files", ".xlsx"), ("All Files", "*.*")))

    #Scan's Document
    df = pd.read_excel(filePath)

    #Reorders Columns to Match Form
    df2 = pd.DataFrame(df, columns = columnList)

    #Fixes Empty Columns and corrects the "Leaking" header
    dfNaN = df2.replace(np.nan, '', regex=True)
    dfNaN.rename(columns={'Repair Made?' : 'Leaking?'}, inplace=True)

    #Exports original data and reformatted data to the original sheet
    with pd.ExcelWriter(filePath, date_format = 'mm/dd/yy', datetime_format='mm/dd/yy') as writer:
        dfNaN.to_excel(writer, sheet_name='IRR Reformatted')
        df.to_excel(writer, sheet_name=('Original Data'))
        

#----------------Main Code Body ----------#
#Creates a Window to Select the Excel File
root = Tk()
root.withdraw()

try:
    ExcelParse(columnList)

#---------------Error Catching------------#
except(FileNotFoundError):
    print('Script Cancelled...\nClosing...')
except:
    messagebox.showerror(title="Error", message= "An error has occured, it is likely a valid excel file was not selected. Please click 'OK' and try again.\n\nIf you continue to recieve this error, please click cancel in the file select window and contact Daniel Steinhauser")
    ExcelParse(columnList)
    


