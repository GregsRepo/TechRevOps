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
FILENAME = "VFX3_IHC_Report.txt"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/'
EXCEL = 'VFX3_IHC_Report.xlsx'

###########################################################################
def vfx3_ihc_report():
    
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVFX3"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/chkRFBSK_F").selected = 'true'
        session.findById("wnd[0]/usr/chkRFBSK_G").selected = 'true'
        session.findById("wnd[0]/usr/chkRFBSK_K").selected = 'true'
        session.findById("wnd[0]/usr/ctxtVKORG").text = "I001"
        session.findById("wnd[0]/usr/txtERNAM-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = start_date
        session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = end_date
        session.findById("wnd[0]/usr/chkRFBSK_K").setFocus()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME 
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (16)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
    except:
        print('No data found for specified date range')

    #load_to_pandas()
##############################################################################

def load_to_pandas():
    VFX3_IHC_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
    
    cols = [c for c in VFX3_IHC_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    VFX3_IHC_DF = VFX3_IHC_DF[cols]
    VFX3_IHC_DF.drop([0, 0], inplace=True) # Drop first empty rows


    VFX3_IHC_DF.to_excel(DIR + EXCEL, index=False)
    print(VFX3_IHC_DF.head())

##############################################################################
if __name__ == "__main__":
    vfx3_ihc_report()
    #load_to_pandas()