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
FILENAME = "PH_Status_Report.txt"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/'
EXCEL = "PH_Status_Report.xlsx"

###########################################################################
def ph_status_report():
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ_PH_RPT"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtS_DATE-LOW").text = start_date
    session.findById("wnd[0]/usr/ctxtS_DATE-HIGH").text = end_date
    session.findById("wnd[0]/usr/chkP_ALL").setFocus()
    session.findById("wnd[0]/usr/chkP_ALL").selected = 'false'
    session.findById("wnd[0]/usr/chkP_NW").setFocus()
    session.findById("wnd[0]/usr/chkP_NW").selected = 'true'
    session.findById("wnd[0]/usr/chkP_BC").setFocus()
    session.findById("wnd[0]/usr/chkP_BC").selected = 'true'
    session.findById("wnd[0]/usr/chkP_PIP").setFocus()
    session.findById("wnd[0]/usr/chkP_PIP").selected = 'true'
    session.findById("wnd[0]/usr/chkP_PE").setFocus()
    session.findById("wnd[0]/usr/chkP_PE").selected = 'true'
    session.findById("wnd[0]/usr/chkP_REP").selected = 'true'
    session.findById("wnd[0]/usr/chkP_REP").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (16)
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    load_to_pandas()
##############################################################################

def load_to_pandas():
    
    PH_STATUS_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
    
    cols = [c for c in PH_STATUS_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    PH_STATUS_DF = PH_STATUS_DF[cols]
    PH_STATUS_DF.drop([0, 0], inplace=True) # Drop first empty rows


    PH_STATUS_DF.to_excel(DIR + EXCEL, index=False)
    print(PH_STATUS_DF.head())

##############################################################################
if __name__ == "__main__":
    #ph_status_report()
    load_to_pandas()