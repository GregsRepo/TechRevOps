###########################################################################
import win32com.client 
import datetime as dt

###########################################################################
SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
session = SapGui.FindById("ses[0]")  

# Get todays date as a datetime object for use in billing report function
today = dt.date.today()
look_back = today - dt.timedelta(days=90)
start_date = look_back.strftime('%m/%d/%Y') 
end_date = today.strftime('%m/%d/%Y') 

###########################################################################
def zaci(REGION, FILEPATH):

    print('Running ZACI ' + REGION +  ' Report...')

    for i in range(2):
        # Set dates and filename for each run of loop
        if i==0:
            minus_ninety_days = today - dt.timedelta(days=90)
            minus_forty_six_days = today - dt.timedelta(days=46)
            start_date = minus_ninety_days.strftime('%m/%d/%Y') 
            end_date = minus_forty_six_days.strftime('%m/%d/%Y')
            FILENAME = 'ZACI_Report_' + REGION + '_1st_half.txt'
        elif i==1:
            look_back = today - dt.timedelta(days=45)
            start_date = look_back.strftime('%m/%d/%Y') 
            end_date = today.strftime('%m/%d/%Y') 
            FILENAME = 'ZACI_Report_' + REGION + '_2nd_half.txt'

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZACI_BILLING_REPORT"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = REGION
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


###########################################################################
def ph_aging(FILEPATH):

    print('Running PH Aging Report...')

    FILENAME = "PH_Aging_Report.txt"
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZ_PH_AGING"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtS_DT_ZAV-LOW").text = start_date
    session.findById("wnd[0]/usr/ctxtS_DT_ZAV-HIGH").text = end_date
    session.findById("wnd[0]/usr/chkP_ALL").setFocus()
    session.findById("wnd[0]/usr/chkP_ALL").selected = 'true'
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (19)
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    return FILENAME 

###########################################################################
def ph_status(FILEPATH):
    
    print('Running PH Status Report...')

    FILENAME = "PH_Status_Report.txt"

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

    return FILENAME

###########################################################################
def vfx3(report, FILEPATH):
    
    if report == 'I001':
        FILENAME = 'VFX3_IHC_Report.txt'
        print('Running VFX3 IHC Report...')
    if report == 'D001':
        FILENAME = 'VFX3_ADIR_Report.txt'
        print('Running VFX3 ADIR Report...')
    if report == '0001':
        FILENAME = 'VFX3_ADUS_Report.txt'
        print('Running VFX3 ADUS Report...')
    
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVFX3"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/chkRFBSK_F").selected = 'true'
        session.findById("wnd[0]/usr/chkRFBSK_G").selected = 'true'
        session.findById("wnd[0]/usr/chkRFBSK_K").selected = 'true'
        session.findById("wnd[0]/usr/ctxtVKORG").text = report 
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
        # Write an empty file so that old data is overwritten
        with open(FILEPATH + '/' + FILENAME, 'w') as empty_csv:
            print(empty_csv)
            pass
    
    return FILENAME

##############################################################################
def bart(report, FILEPATH):

    if report == '1':
        FILENAME = "BART_Error_Report.txt"
        print('Running Bart Error Report...')
    if report == '3':
        FILENAME = 'BART_Duplicate_Report.txt'
        print('Running Bart Duplicate Report...')
    if report == '7':
        FILENAME = 'BART_No_Provisioning.txt'
        print('Running Bart No Provisioning Report...')

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
    session.findById("wnd[0]/usr/ctxtP_AFDATE").text = start_date
    session.findById("wnd[0]/usr/ctxtP_ATDATE").text = end_date
    session.findById("wnd[0]/usr/ctxtP_STATUS").text = report 
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

    return FILENAME

###########################################################################
def v_uc(FILEPATH):

    print('Running V_UC Report...')

    FILENAME = 'V_UC_Report.txt'
    
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
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (20)
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    return FILENAME

###########################################################################
def zisxerror(FILEPATH):

    print('Running ZISXERROR Report...')

    FILENAME = 'ZISXERROR_Report.txt'
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZISXERROR"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtI15-LOW").text = start_date
    session.findById("wnd[0]/usr/ctxtI15-HIGH").text = end_date
    session.findById("wnd[0]/usr/ctxtLIST_BRE").text = "999"
    session.findById("wnd[0]/usr/txtMAX_SEL").text = "999 "
    session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
    session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = (10)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/mbar/menu[1]/menu[5]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = FILEPATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FILENAME
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (20)
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    return FILENAME

###########################################################################
        