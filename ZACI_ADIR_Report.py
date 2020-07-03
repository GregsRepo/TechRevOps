###########################################################################
import win32com.client 
import pandas as pd
from pandas import ExcelWriter
import datetime as dt

###########################################################################
SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
session = SapGui.FindById("ses[0]")  

# Get todays date as a datetime object for use in billing report function
today = dt.date.today()

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
            minus_ninety_days = today - dt.timedelta(days=90)
            minus_forty_six_days = today - dt.timedelta(days=46)
            start_date = minus_ninety_days.strftime('%m/%d/%Y') 
            end_date = minus_forty_six_days.strftime('%m/%d/%Y')
            FILENAME = "ZACI_Report_ADIR_1st_half.txt"
        elif i==1:
            look_back = today - dt.timedelta(days=45)
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


    load_to_pandas()
##############################################################################

def load_to_pandas():
    FILENAME_1 = "ZACI_Report_ADIR_1st_half.txt"
    FILENAME_2 = "ZACI_Report_ADIR_2nd_half.txt"

    # read the first file into seperate dataframes twice. 
    # This is needed because of the way the data is stacked in the .txt file
    # further down we then read each alternating line into seperate dataframes 
    # and append the dataframes into 1 df to give us legible data
    df1 = pd.read_csv(FILEPATH + '/' + FILENAME_1, skiprows=3, sep='|', engine='python')
    df2 = pd.read_csv(FILEPATH + '/' + FILENAME_1, skiprows=3, sep='|', engine='python')
    
    
    # drop the empty 'unnamed' columns
    cols1 = [c for c in df1.columns if c.lower()[:7] != 'unnamed'] 
    cols2 = [c for c in df2.columns if c.lower()[:7] != 'unnamed']
    df1 = df1[cols1]
    df2 = df2[cols2]


    # # drop the the first 2 rows from df1 and retrieve every 2nd row and then reset index
    df1 = df1.drop([0, 1])
    df1 = df1.iloc[::2]
    df1 = df1.reset_index(drop=True)

    # drop the the first 2 rows from df2 and retrieve every 2nd row and then reset index
    df2 = df2.drop([1, 2])
    df2.columns = df2.iloc[0]
    df2 = df2.iloc[1::2]
    df2 = df2.reset_index(drop=True)
    

    # concatenate dataframe 1 & 2 into one dataframe
    first_half = pd.concat([df1, df2], axis=1, sort=False)
    

    # repeat the code above for file 2
    df3 = pd.read_csv(FILEPATH + '/' + FILENAME_2, skiprows=3, sep='|', engine='python')
    df4 = pd.read_csv(FILEPATH + '/' + FILENAME_2, skiprows=3, sep='|', engine='python')
    cols3 = [c for c in df3.columns if c.lower()[:7] != 'unnamed'] 
    cols4 = [c for c in df4.columns if c.lower()[:7] != 'unnamed']
    df3 = df3[cols3]
    df4 = df4[cols4]
    df3 = df3.drop([0, 1])
    df3 = df3.iloc[::2]
    df3 = df3.reset_index(drop=True)
    df4 = df4.drop([1, 2])
    df4.columns = df4.iloc[0]
    df4 = df4.iloc[1::2]
    df4 = df4.reset_index(drop=True)


    # concatenate dataframe 3 & 4 into one dataframe
    second_half = pd.concat([df3, df4], axis=1, sort=False)

    # clean the new dataframes by dropping empty columns and rows, and stripping whitespcae from the column names
    first_half.dropna(axis = 1, how ='all', inplace = True) 
    second_half.dropna(axis = 1, how ='all', inplace = True)
    first_half.dropna(axis = 0, how ='all', inplace = True) 
    second_half.dropna(axis = 0, how ='all', inplace = True)
    first_half.columns = first_half.columns.str.strip()
    second_half.columns = second_half.columns.str.strip()

    # concatenate dataframe 1 & 2 into one ZACI ADIR Report
    ZACI_ADIR = pd.concat([first_half, second_half], sort=False)

    # output to excel
    ZACI_ADIR.to_excel(DIR + EXCEL, index=False)

    print(ZACI_ADIR.shape)

##############################################################################
if __name__ == "__main__":
    zaci_billing_report()
    #load_to_pandas()