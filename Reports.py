
###########################################################################################
def zaci_adir(session):

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZACI_BILLING_REPORT"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "ADIR"
    session.findById("wnd[0]/usr/ctxtS_SUBPRO-LOW").text = "ZCTR"
    session.findById("wnd[0]/usr/ctxtS_SUBPRO-HIGH").text = "ZSUB"
    session.findById("wnd[0]/usr/ctxtS_DT_O-LOW").text = "04/01/2020"
    session.findById("wnd[0]/usr/ctxtS_DT_O-HIGH").text = "05/28/2020"
    session.findById("wnd[0]/usr/chkP_BILL").setFocus()
    session.findById("wnd[0]/usr/chkP_BILL").selected = "true"
    session.findById("wnd[0]/usr/chkP_BILL_N").setFocus()
    session.findById("wnd[0]/usr/chkP_BILL_N").selected = "true"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:/Users/grwillia/OneDrive - Adobe Systems Inc/Backup/Desktop/SAP_Reports"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZACI_Report_ADIR.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

###########################################################################################
def bart_duplicate(session):

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZRPT"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/btnSD_BUTTON").press()
    session.findById("wnd[0]/usr/btnEO_BUTTON").press()
    session.findById("wnd[0]/usr/sub:ZAPMZRPT:9100/radG20_TPROG-RADIO[3,0]").select()
    session.findById("wnd[0]/usr/sub:ZAPMZRPT:9100/radG20_TPROG-RADIO[3,0]").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtP_AFDATE").text = "05/21/2020"
    session.findById("wnd[0]/usr/ctxtP_ATDATE").text = "05/30/2020"
    session.findById("wnd[0]/usr/txtP_NAME").text = "PROFSVCS_CIC"
    session.findById("wnd[0]/usr/txtP_NAME").setFocus()
    session.findById("wnd[0]/usr/txtP_NAME").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlALVCDRGRID/shellcont/shell").pressToolbarContextButton = "&MB_EXPORT" # Added '=' sign. need to see if this works
    session.findById("wnd[0]/usr/cntlALVCDRGRID/shellcont/shell").selectContextMenuItem = "&PC"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "/du1isi0/order_management/TechRevOps/Reports/SAP_Reports/Report Downloads"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "BART_REPORT.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
    session.findById("wnd[1]/tbar[0]/btn[11]").press()