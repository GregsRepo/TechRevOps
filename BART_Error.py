###########################################################################
import win32com.client 
import pandas as pd
from pandas import ExcelWriter
import datetime as dt
import Reports

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
FILENAME = "BART_Error_Report.txt"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/'
EXCEL = 'BART_Error_Report.xlsx'

###########################################################################
def bart_error_report():
    
   
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZRPT"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/btnSD_BUTTON").press()
    session.findById("wnd[0]/usr/btnEO_BUTTON").press()
    session.findById("wnd[0]/usr/sub:ZAPMZRPT:9100/radG20_TPROG-RADIO[3,0]").select()
    session.findById("wnd[0]/usr/sub:ZAPMZRPT:9100/radG20_TPROG-RADIO[3,0]").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = (2)
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    session.findById("wnd[0]/usr/ctxtP_AFDATE").text = "05/01/2019"
    session.findById("wnd[0]/usr/ctxtP_ATDATE").text = "06/28/2019"
    session.findById("wnd[0]/usr/ctxtP_STATUS").text = "1"
    session.findById("wnd[0]/usr/ctxtP_ATDATE").setFocus()
    session.findById("wnd[0]/usr/ctxtP_ATDATE").caretPosition = (10)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlALVCDRGRID/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlALVCDRGRID/shellcont/shell").selectContextMenuItem ("&PC")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    

    #load_to_pandas()
##############################################################################

def load_to_pandas():
    BART_ERROR_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=4, sep='|', engine='python')
    
    cols = [c for c in BART_ERROR_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    BART_ERROR_DF = BART_ERROR_DF[cols]
    BART_ERROR_DF.drop([0, 0], inplace=True) # Drop first empty rows


    BART_ERROR_DF.to_excel(DIR + EXCEL, index=False)
    print(BART_ERROR_DF.head())

##############################################################################
if __name__ == "__main__":
    #bart_error_report()
    load_to_pandas()