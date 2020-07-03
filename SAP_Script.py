###########################################################################
import win32com.client 
import pandas as pd
from pandas import ExcelWriter
import datetime as dt

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
FILENAME = "ZACI_Report_ADIR.txt"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/'
EXCEL = "ZACI_Report_ADIR.xlsx"

###########################################################################
def zaci_billing_report():
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZACI_BILLING_REPORT"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "ADIR"
    session.findById("wnd[0]/usr/ctxtS_SUBPRO-LOW").text = "ZCTR"
    session.findById("wnd[0]/usr/ctxtS_SUBPRO-HIGH").text = "ZSUB"
    session.findById("wnd[0]/usr/ctxtS_DT_O-LOW").text = start_date
    session.findById("wnd[0]/usr/ctxtS_DT_O-HIGH").text = end_date
    session.findById("wnd[0]/usr/chkP_BILL").setFocus()
    session.findById("wnd[0]/usr/chkP_BILL").selected = "true"
    session.findById("wnd[0]/usr/chkP_BILL_N").setFocus
    session.findById("wnd[0]/usr/chkP_BILL_N").selected = "true"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (16)
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    #load_to_pandas()
##############################################################################

def load_to_pandas():
    ZACI_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
    
    ZACI_DF.drop([0, 1], inplace=True) # Drop first empty rows
    ZACI_DF = ZACI_DF[:-1] # Drop last row containing dash's --
    ZACI_DF.dropna(axis=1, how='all', inplace=True) # Drop Nan columns



    ZACI_DF.to_excel(DIR + EXCEL, index=False)
    print(ZACI_DF.head())

##############################################################################
if __name__ == "__main__":
    #zaci_billing_report()
    load_to_pandas()