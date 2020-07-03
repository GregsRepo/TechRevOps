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
# look_back = today - dt.timedelta(days=45)
# start_date = look_back.strftime('%m/%d/%Y') 
# end_date = today.strftime('%m/%d/%Y') 

# Set the filepath and filename for the downloaded report
#FILEPATH = "//du1isi0/order_management/TechRevOps/Reports/SAP_Reports/Report Downloads"
FILEPATH = "C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports"


# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/'
EXCEL = "ZACI_Report_ADIR.xlsx"

###########################################################################
def zaci_billing_report():

    
    for i in range(2):
        # Set dates and filename for each run of loop
        if i==0:
            ninety_days = today - dt.timedelta(days=30)
            forty_six_days = today - dt.timedelta(days=16)
            start_date = ninety_days.strftime('%m/%d/%Y') 
            end_date = forty_six_days.strftime('%m/%d/%Y')
            FILENAME = "ZACI_Report_ADIR_1st_half.txt"
        elif i==1:
            look_back = today - dt.timedelta(days=15)
            start_date = look_back.strftime('%m/%d/%Y') 
            end_date = today.strftime('%m/%d/%Y') 
            FILENAME = "ZACI_Report_ADIR_2nd_half.txt"

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
    # ZACI_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
    
    # ZACI_DF.drop([0, 1], inplace=True) # Drop first empty rows
    # ZACI_DF = ZACI_DF[:-1] # Drop last row containing dash's --
    # ZACI_DF.dropna(axis=1, how='all', inplace=True) # Drop Nan columns
    FILENAME_1 = "ZACI_Report_ADIR_1st_half.txt"
    FILENAME_2 = "ZACI_Report_ADIR_2nd_half.txt"

    df1 = pd.read_csv(FILEPATH + '/' + FILENAME_1, skiprows=3, sep='|', engine='python')
    df2 = pd.read_csv(FILEPATH + '/' + FILENAME_2, skiprows=3, sep='|', engine='python')
    
    # l
    cols1 = [c for c in df1.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    cols2 = [c for c in df2.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    df1 = df1[cols1]
    df2 = df2[cols2]

    #drop the the first 2 rows, the first column named 'F1' and retrieve every 2nd row
    df1 = df1.drop([0, 1])
    #df1 = df1.drop(columns=['F1'])
    df1 = df1.iloc[::2]

    #drop the the first 2 rows, and retrieve every 2nd row
    df2 = df2.drop([0, 1])
    df2 = df2.iloc[::2]

    #concatenate dataframe 1 & 2 into one ZACI ADIR Report
    ZACI_ADIR = pd.concat([df1, df2], axis=1, sort=False)

    ZACI_ADIR.to_excel(DIR + EXCEL, index=False)
    print(ZACI_ADIR.head())

     
     
     '''Output'''
    #ZACI_ADUS.to_excel(DIR + EXCEL, index=False)
    # first_half.to_excel(DIR + 'adus1sthalf.xlsx', index=False)
    # second_half.to_excel(DIR + 'adus2ndhalf.xlsx', index=False)
    # df1.to_excel(DIR + 'df1.xlsx', index=False)
    # df2.to_excel(DIR + 'df2.xlsx', index=False)
    # df3.to_excel(DIR + 'df3.xlsx', index=False)
    # df4.to_excel(DIR + 'df4.xlsx', index=False)

    '''Print'''
    #print(ZACI_ADUS.head())
    
    # print('First Half', first_half.shape)
    # print('Second Half', second_half.shape)
    # print('ZACI Report', ZACI_ADUS.shape)

    
    
    # firsthalf = list(first_half)
    # secondhalf = list(second_half)

    # diff = list(set(second_half) - set(first_half))

    # print('\n\n', diff)

    # print('\n\n')
    # print(firsthalf)
    # print('\n\n')
    # print(secondhalf)

##############################################################################
if __name__ == "__main__":
    #zaci_billing_report()
    load_to_pandas()