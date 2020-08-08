###########################################################################
import win32com.client 
import pandas as pd
from pandas import ExcelWriter

###########################################################################
SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
session = SapGui.FindById("ses[0]")  

FILEPATH = "C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports/"

FILENAME = "ZACI_Test.txt"

def zaci_billing_report():

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZACI_BILLING_REPORT"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "ADUS"
        session.findById("wnd[0]/usr/ctxtS_SUBPRO-LOW").text = "ZCTR"
        session.findById("wnd[0]/usr/ctxtS_SUBPRO-HIGH").text = "ZSUB"
        session.findById("wnd[0]/usr/ctxtS_DT_O-LOW").text = '05/01/2020'
        session.findById("wnd[0]/usr/ctxtS_DT_O-HIGH").text = '06/30/2020'
        session.findById("wnd[0]/usr/chkP_BILL").setFocus()
        session.findById("wnd[0]/usr/chkP_BILL").selected = "true"
        session.findById("wnd[0]/usr/chkP_BILL_N").setFocus
        session.findById("wnd[0]/usr/chkP_BILL_N").selected = "true"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (16)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #load_pandas(FILEPATH, FILENAME)

def load_pandas(FILEPATH, FILENAME):
    ZACI_DF = pd.read_csv(FILEPATH + FILENAME, skiprows=3, sep='\t', engine='python')

    cols = [c for c in ZACI_DF.columns if c.lower()[:7] != 'unnamed'] 
    ZACI_DF = ZACI_DF[cols]

    print(ZACI_DF.head)

    ZACI_DF.dropna(how='all')
    ZACI_DF.to_excel(FILEPATH + 'ZACI_Test.xlsx', index=False)

if __name__ == "__main__":
    load_pandas(FILEPATH, FILENAME)