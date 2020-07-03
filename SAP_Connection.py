# Importing the Libraries
#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------
import sys, win32com.client
import Reports

fpath = "\\\\du1isi0\\order_management\\TechRevOps\\Reports\\SAP_Reports\\Report Downloads"
ffilename = "BART_Error_Report.txt"

col1 = '12\\01\\2020'
col3 = '02\\28\\2020'
#-Sub Main--------------------------------------------------------------
def Main():

  try:

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
      return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
      SapGuiAuto = None
      return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
      application = None
      SapGuiAuto = None
      return

    session = connection.Children(1)
    if not type(session) == win32com.client.CDispatch:
      connection = None
      application = None
      SapGuiAuto = None
      return

    Reports.zaci_adir(session)

    #Rem ADDED BY EXCEL *************************************
    # session.findById("wnd[0]").maximize()
    # session.findById("wnd[0]/tbar[0]/okcd").text = "/nZACI_BILLING_REPORT"
    # session.findById("wnd[0]").sendVKey (0)
    # session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "ADIR"
    # session.findById("wnd[0]/usr/ctxtS_SUBPRO-LOW").text = "ZCTR"
    # session.findById("wnd[0]/usr/ctxtS_SUBPRO-HIGH").text = "ZSUB"
    # session.findById("wnd[0]/usr/ctxtS_DT_O-LOW").text = "04/01/2019"
    # session.findById("wnd[0]/usr/ctxtS_DT_O-HIGH").text = "06/30/2019"
    # session.findById("wnd[0]/usr/chkP_BILL").setFocus()
    # session.findById("wnd[0]/usr/chkP_BILL").selected = "true"
    # session.findById("wnd[0]/usr/chkP_BILL_N").setFocus()
    # session.findById("wnd[0]/usr/chkP_BILL_N").selected = "true"
    # session.findById("wnd[0]/tbar[1]/btn[8]").press()
    # session.findById("wnd[0]/tbar[1]/btn[45]").press()
    # session.findById("wnd[1]/tbar[0]/btn[0]").press()
    #Rem FINALIZATION CONTROL CHECK ************************

  except:
    print(sys.exc_info()[0])

  finally:
    session = None
    connection = None
    application = None
    SapGuiAuto = None

#-Main------------------------------------------------------------------
if __name__ == "__main__":
  Main()

#-End-------------------------------------------------------------------