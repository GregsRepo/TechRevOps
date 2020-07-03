###########################################################################
import win32com.client 
import pandas as pd
import numpy as np
from pandas import ExcelWriter
import datetime as dt
import csv

###########################################################################
SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
session = SapGui.FindById("ses[0]")  

# Datetime formatting
# Get todays date and the number of days to look back and format them to strings
today = dt.date.today()
look_back = today - dt.timedelta(days=30)
start_date = look_back.strftime('%m/%d/%Y') 
end_date = today.strftime('%m/%d/%Y') 

# Set the filepath and filename for the downloaded report
#FILEPATH = "//du1isi0/order_management/TechRevOps/Reports/SAP_Reports/Report Downloads"
FILEPATH = "C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports"
FILENAME = "V_UC_Report_formatted.txt"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/'
EXCEL = "V_UC_Report2.xlsx"

###########################################################################
def v_uc_report():
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nV_UC"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = (0)
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = (3)
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (11)
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    #load_to_pandas()
##############################################################################

def load_to_pandas():
    
    # Read V_UC Report into pandas dataframe
    V_UC_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=2, sep='\t', engine='python') # 
    
    # Drop the empty unnamed columns
    cols = [c for c in V_UC_DF.columns if c.lower()[:7] != 'unnamed'] 
    V_UC_DF = V_UC_DF[cols]

    # Strip whitespace from column names
    V_UC_DF.rename(columns=lambda x: x.strip(), inplace=True)

    # Rename Item column to Doc Number
    V_UC_DF.rename(columns={'Item':'Doc Number'}, inplace=True)

    # Create new column 'Item Number' to store 
    V_UC_DF['Item No.'] = np.where(V_UC_DF['Doc Number']>=50, '', V_UC_DF['Doc Number'])

    # Convert Item No. from object to float
    V_UC_DF['Item No.'] = pd.to_numeric(V_UC_DF['Item No.'], errors='coerce')
    
    # Create new column 'Doc Type' to record whether the document is a ZAV or an Order
    # If > 200M 'ZAV, if < 100M ' Order, else ' ' 
    V_UC_DF['Doc Type'] = np.where((V_UC_DF['Doc Number'] > 200000000) , 'ZAV',   
                 np.where((V_UC_DF['Doc Number'] < 100000000) , '', 
                  np.where((V_UC_DF['Doc Number'] > 100000000) & (V_UC_DF['Doc Number'] < 200000000), 'Order',  
                    '')))   
    
    # Replace item numbers with None in Doc Number column
    V_UC_DF['Doc Number'] = V_UC_DF['Doc Number'].where(V_UC_DF['Doc Number'] > 1000, None) 

    # Reorder the columns
    V_UC_DF = V_UC_DF[['Doc Number','Doc Type','Item No.','Short Description','General', 'Delivery', 'BillingDoc', 'Price', \
        'Goods mov.', 'Pck/putaw.', 'Pack']]
    
    V_UC_DF.to_excel(DIR + EXCEL, index=False)
    print(V_UC_DF.head())
    print(V_UC_DF['Item No.'].dtype)


##############################################################################
if __name__ == "__main__":
    #v_uc_report()
    load_to_pandas()