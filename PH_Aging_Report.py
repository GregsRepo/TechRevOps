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
FILENAME = "PH_Aging_Report.txt"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/'
EXCEL = "PH_Aging_Report.xlsx"

###########################################################################
def ph_aging_report():
    
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

    #load_to_pandas()
##############################################################################

def ph_aging_dataframe():
    
    PH_AGING = FILEPATH + '/' + FILENAME

    # Read PH Aging into dataframe
    PH_AGING_DF = pd.read_csv(PH_AGING, skiprows=3, sep='|', engine='python')

    # Drop the empty unnamed columns from PH Aging dataframe
    cols = [c for c in PH_AGING_DF.columns if c.lower()[:7] != 'unnamed'] 
    PH_AGING_DF = PH_AGING_DF[cols]
    PH_AGING_DF.drop([0, 0], inplace=True) # Drop first empty rows

    # Rename columns
    PH_AGING_DF.columns = ['Opportunity ID', 'Region', 'DR Number', 'Sales Org', 'EU Country Code', 'Currency', 'Customer', 'End User',
                            'EU Cust Name', 'Sales Order Created By', 'Document No.', 'Sales Doc.', 'Doc Type', 'Amount', 'After PC', 
                            'New', 'Booking Complete', 'Provisioning in Progress', 'Provisioning Completed', 'Provisioning Error',
                            'Total No. of Days', 'Create Date(ZCC)', 'Create Date(ZAV)', 'Last Status Date']

    # Join data taken from PH Status Report and combine it wiht DF to create the PH Aging Report
    PH_AGING_DF = pd.merge(PH_AGING_DF, JOIN)

    # Add a Notes column at index 0
    PH_AGING_DF.insert(loc=0, column='Notes', value='')

    # Trim whitespace from all cells
    PH_AGING_DF = PH_AGING_DF.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # Filter out rejected orders. Used 'ject' becuase we can have mixture of lower/upper case as well as ZAV Rejected
    PH_AGING_DF = PH_AGING_DF[~PH_AGING_DF['Opportunity ID'].str.contains("ject", na=False)]
    PH_AGING_DF = PH_AGING_DF[~PH_AGING_DF['Opportunity ID'].str.contains("JECT", na=False)]

    # Change these columns to datetime objects
    today = pd.Timestamp('today').floor('D')
    PH_AGING_DF['Create Date(ZCC)'] = pd.to_datetime(PH_AGING_DF['Create Date(ZCC)'])
    PH_AGING_DF['Contract Start Date'] = pd.to_datetime(PH_AGING_DF['Contract Start Date'])
    PH_AGING_DF['Create Date(ZAV)'] = pd.to_datetime(PH_AGING_DF['Create Date(ZAV)'])
    
    # Create dataframe for ZAV User Status equals New, Booking Complete, Prov in Porgress, or Prov Error orders
    NEW = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] == 'New']
    BOOKING_COMPLETE = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] == 'Booking Complete']
    PROV_IN_PROGRESS = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] == 'Provisioning in Progress']
    PROVIONING_ERROR = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] == 'Provisioning Error']

    # Add notes to Booking Complete dataframe
    BOOKING_COMPLETE['Notes'] = np.where((BOOKING_COMPLETE['Create Date(ZAV)'] == today) , "Created Today",  
                                np.where((BOOKING_COMPLETE['Contract Start Date'] == today) , "Created Today",
                                np.where((BOOKING_COMPLETE['Create Date(ZAV)'] > today) , "Future Start Date",  
                                np.where((BOOKING_COMPLETE['Contract Start Date'] > today) , "Future Start Date",
                                np.where((BOOKING_COMPLETE['Header Block'] == 'ZH : Waiting on PO') , 'On Header Block',
                                np.where((BOOKING_COMPLETE['Header Block'] == 'PP : Provisioning Pending') , 'On Header Block', 
                                '')))))) 

    # Add notes to PROV_IN_PROGRESS dataframe
    PROV_IN_PROGRESS['Notes'] = np.where((PROV_IN_PROGRESS['Create Date(ZAV)'] == today) , "Created Today",  
                    np.where((PROV_IN_PROGRESS['Contract Start Date'] == today) , "Created Today",
                    np.where((PROV_IN_PROGRESS['Create Date(ZAV)'] > today) , "Future Start Date",  
                    np.where((PROV_IN_PROGRESS['Contract Start Date'] > today) , "Future Start Date",
                    np.where((PROV_IN_PROGRESS['Header Block'] == 'ZH : Waiting on PO') , 'On Header Block',
                    np.where((PROV_IN_PROGRESS['Header Block'] == 'PP : Provisioning Pending') , 'On Header Block', 
                    '')))))) 
    

    # Drop anything that is New and Create Date is less than today
    PH_AGING_DF.drop(PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'New')  & (PH_AGING_DF['Create Date(ZCC)'].dt.date < today)].index, inplace=True) 
    
    # Drop anything that is Booking Complete and Waiting on PO or Provisioning Pending
    PH_AGING_DF.drop(PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'Booking Complete')  & (PH_AGING_DF['Header Block'] != 'ZH : Waiting on PO')].index, inplace=True) 
    PH_AGING_DF.drop(PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'Booking Complete')  & (PH_AGING_DF['Header Block'] != 'PP : Provisioning Pending')].index, inplace=True) 
    
    # Drop anything that is Provisioning in Progress and Create Date less than today and Contract Start date less than today 
    PH_AGING_DF.drop(PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'Provisioning in Progress')  & (PH_AGING_DF['Create Date(ZCC)'].dt.date < today)].index, inplace=True) 
    PH_AGING_DF.drop(PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'Provisioning in Progress')  & (PH_AGING_DF['Contract Start Date'] < today)].index, inplace=True) 
    
    '''Find out why we do this? If its necessary, create the sales doc df in PH_Status and pass it in for merge'''
    #PH_AGING_DF = pd.merge(SALES_DOC, PH_AGING_DF, on='Sales Doc.')

    # Add a note for anything that is Waiting on PO or has a Future Start Date
    PH_AGING_DF['Notes'] = np.where((PH_AGING_DF['Header Block'] == "ZH : Waiting on PO") , "Billing Block",   
                 np.where((PH_AGING_DF['Contract Start Date'].dt.date > today) , 'Future Start Date',
                  'Review'))   
    

    # Add a Notes column at index 0
    PH_AGING_DF.insert(loc=0, column='Review Comments', value='')

    '''Might use this at a later date'''
    # # This section reads in the notes from previous review file using Sales Doc as identifier
    # AGING_COMMENTS = pd.read_excel(DIR + EXCEL, sheet_name='Aging Review') 
    # AGING_COMMENTS = AGING_COMMENTS[['Review Comments', 'Sales Doc.']] # Take only the Notes and Source Transaction Id column
    # PH_AGING_DF = pd.merge(PH_AGING_DF, AGING_COMMENTS, on='Sales Doc.', how='left') # Merge previous notes with new dataframe based on Source Transaction ID
    

    return BOOKING_COMPLETE, PROV_IN_PROGRESS, NEW, PROVIONING_ERROR, PH_AGING_DF

##############################################################################
if __name__ == "__main__":
    #ph_aging_report()
    load_to_pandas()