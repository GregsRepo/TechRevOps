###########################################################################
import win32com.client 
import numpy as np
import pandas as pd
from pandas import ExcelWriter
import datetime as dt


###########################################################################
SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
session = SapGui.FindById("ses[0]")  

# Get todays date as a datetime object for use in billing report function
today = dt.date.today()
#todays_date = today.strftime('%m/%d/%Y') 
todays_date = '06/12/2020'

# Set the filepath and filename for the downloaded report
#FILEPATH = "//du1isi0/order_management/TechRevOps/Reports/SAP_Reports/Report Downloads"
FILEPATH = "C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/ZACI/'
#EXCEL = "ZACI_Report_ADUS.xlsx"
EXCEL = "ZACI_Report.xlsx"

###########################################################################
def zaci_billing_report():

    '''Change this to a class and pass in Report type as 'region' (i.e. ADIR or ADUS)'''
    for i in range(2):
        # Set dates and filename for each run of loop
        if i==0:
            minus_ninety_days = today - dt.timedelta(days=90)
            minus_forty_six_days = today - dt.timedelta(days=46)
            start_date = minus_ninety_days.strftime('%m/%d/%Y') 
            end_date = minus_forty_six_days.strftime('%m/%d/%Y')
            FILENAME = "ZACI_Report_ADUS_1st_half.txt"
        elif i==1:
            look_back = today - dt.timedelta(days=45)
            start_date = look_back.strftime('%m/%d/%Y') 
            end_date = today.strftime('%m/%d/%Y') 
            FILENAME = "ZACI_Report_ADUS_2nd_half.txt"

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZACI_BILLING_REPORT"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "ADUS"
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

    '''This needs to be changed to a nested loop instead of duplicating code
    First loop will read in which half to work on and second loop will split the file, 
    format the seperate pieces, and concatenate back to one dataframe. 
    Make this a class and pass in the Region (i.e AIDR or ADUS)'''
    
    FILENAME_1 = "ZACI_Report_ADUS_1st_half.txt"
    FILENAME_2 = "ZACI_Report_ADUS_2nd_half.txt"

    # read the first file into seperate dataframes twice. 
    # This is needed because of the way the data is stacked in the .txt file
    # further down we then read each alternating line into seperate dataframes 
    # and append the dataframes into 1 ZACI_ADUS to give us legible data
    ZACI_ADUS1 = pd.read_csv(FILEPATH + '/' + FILENAME_1, skiprows=3, sep='|', engine='python')
    ZACI_ADUS2 = pd.read_csv(FILEPATH + '/' + FILENAME_1, skiprows=3, sep='|', engine='python')
    
    
    # drop the empty 'unnamed' columns
    cols1 = [c for c in ZACI_ADUS1.columns if c.lower()[:7] != 'unnamed'] 
    cols2 = [c for c in ZACI_ADUS2.columns if c.lower()[:7] != 'unnamed']
    ZACI_ADUS1 = ZACI_ADUS1[cols1]
    ZACI_ADUS2 = ZACI_ADUS2[cols2]


    # # drop the the first 2 rows from ZACI_ADUS1 and retrieve every 2nd row and then reset index
    ZACI_ADUS1 = ZACI_ADUS1.drop([0, 1])
    ZACI_ADUS1 = ZACI_ADUS1.iloc[::2]
    ZACI_ADUS1 = ZACI_ADUS1.reset_index(drop=True)

    # drop the the first 2 rows from ZACI_ADUS2 and retrieve every 2nd row and then reset index
    ZACI_ADUS2 = ZACI_ADUS2.drop([1, 2])
    ZACI_ADUS2.columns = ZACI_ADUS2.iloc[0]
    ZACI_ADUS2 = ZACI_ADUS2.iloc[1::2]
    ZACI_ADUS2 = ZACI_ADUS2.reset_index(drop=True)
    

    # concatenate dataframe 1 & 2 into one dataframe
    first_half = pd.concat([ZACI_ADUS1, ZACI_ADUS2], axis=1, sort=False)
    

    # repeat the code above for file 2
    ZACI_ADUS3 = pd.read_csv(FILEPATH + '/' + FILENAME_2, skiprows=3, sep='|', engine='python')
    ZACI_ADUS4 = pd.read_csv(FILEPATH + '/' + FILENAME_2, skiprows=3, sep='|', engine='python')
    cols3 = [c for c in ZACI_ADUS3.columns if c.lower()[:7] != 'unnamed'] 
    cols4 = [c for c in ZACI_ADUS4.columns if c.lower()[:7] != 'unnamed']
    ZACI_ADUS3 = ZACI_ADUS3[cols3]
    ZACI_ADUS4 = ZACI_ADUS4[cols4]
    ZACI_ADUS3 = ZACI_ADUS3.drop([0, 1])
    ZACI_ADUS3 = ZACI_ADUS3.iloc[::2]
    ZACI_ADUS3 = ZACI_ADUS3.reset_index(drop=True)
    ZACI_ADUS4 = ZACI_ADUS4.drop([1, 2])
    ZACI_ADUS4.columns = ZACI_ADUS4.iloc[0]
    ZACI_ADUS4 = ZACI_ADUS4.iloc[1::2]
    ZACI_ADUS4 = ZACI_ADUS4.reset_index(drop=True)


    # concatenate dataframe 3 & 4 into one dataframe
    second_half = pd.concat([ZACI_ADUS3, ZACI_ADUS4], axis=1, sort=False)

    # clean the new dataframes by dropping empty columns and rows, and stripping whitespcae from the column names
    first_half.dropna(axis = 1, how ='all', inplace = True) 
    second_half.dropna(axis = 1, how ='all', inplace = True)
    first_half.dropna(axis = 0, how ='all', inplace = True) 
    second_half.dropna(axis = 0, how ='all', inplace = True)
    first_half.columns = first_half.columns.str.strip()
    second_half.columns = second_half.columns.str.strip()

    # concatenate dataframe 1 & 2 into one ZACI ADUS Report
    ZACI_ADUS = pd.concat([first_half, second_half], sort=False)


    filter_dataframe(ZACI_ADUS)

##############################################################################
def filter_dataframe(ZACI_ADUS):

    '''Change this to take in both dataframe regions and concatenate before performing below operations.
    Will also need to read in DX and DME sheet'''

    # strip whitespace from all columns
    ZACI_ADUS = ZACI_ADUS.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # Filter out deleted orders
    ZACI_ADUS = ZACI_ADUS[ZACI_ADUS['Invoice Order Deleted'] != 'X']

    # Add comments field to store order billing block reason
    ZACI_ADUS['Comments'] = ''
    # Add comments for line items that will not need review based on billing blocks or a Clarification Case number
    ZACI_ADUS['Comments'] = np.where((ZACI_ADUS['Clarification Case Number'] != ''), 'Case ' + ZACI_ADUS['Clarification Case Number'], 
        np.where((ZACI_ADUS['Created On'] == todays_date), 'Created Today',
        np.where((ZACI_ADUS['Billed On'] == todays_date), 'Created Today',
        np.where((ZACI_ADUS['CA Invoice Lock'] == 'Y'), 'Contr. Acc. Lock', 
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block', 
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block',
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block', 
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block', 
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block', 
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block',
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('ZS-')), 'Waiting on PO block', 
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('ZS-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('ZS-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('ZV-')), 'Waiting on PO block', 
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('ZV-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('ZV-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Header Bill Block'].str.contains('15-') ), 'Waiting on PO block',
        np.where((ZACI_ADUS['Item Bill Block'].str.contains('15-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Bill Plan Bill Block'].str.contains('15-')), 'Waiting on PO block',
        np.where((ZACI_ADUS['Header Bill Block'] != ''), 'Review',
        np.where((ZACI_ADUS['Item Bill Block'] != ''), 'Review',
        np.where((ZACI_ADUS['Bill Plan Bill Block'] != ''), 'Review','Review')))))))))))))))))))))))))))))))


    # Convert Source Transaction ID from object to number for the Join below
    ZACI_ADUS['Source Transaction ID'] = pd.to_numeric(ZACI_ADUS['Source Transaction ID']) 
    # Drop duplciate Source Transaction ID's. These only occur on Clarification Cases which we don't need anyway            
    ZACI_ADUS.drop_duplicates(subset ='Source Transaction ID', keep = False, inplace = True) 
    print('ZACI Report', ZACI_ADUS.shape)

    # This section reads in the notes from previous review file using Source Transaction ID as identifier
    DX_NOTES = pd.read_excel(DIR + EXCEL, sheet_name='DX') #'/ZACI/' + 'ZACI_Report.xlsx'
    DME_NOTES = pd.read_excel(DIR + EXCEL, sheet_name='DME') 
    DX_NOTES = DX_NOTES[['Notes', 'Source Transaction ID']] # Take only the Notes and Source Transaction Id column
    DME_NOTES = DME_NOTES[['Notes', 'Source Transaction ID']] # Take only the Notes and Source Transaction Id column
    frames = [DX_NOTES, DME_NOTES]
    NOTES = pd.concat(frames) # Concatenate the Notes dataframe together for merge with ZACI frame
    print('Notes DF', NOTES.shape)
    ZACI_ADUS = pd.merge(ZACI_ADUS, NOTES, on='Source Transaction ID', how='left') # Merge previous notes with new dataframe based on Source Transaction ID
    

    # Reorder columns
    ZACI_ADUS = ZACI_ADUS[[
        'Notes',	'Comments', 'Status',	'Company Code',	'Source Transaction ID',	'Sales Doc Number',	'Contract Item',	
    'Item Number',	'Product',	'Product Description',	'Quantity',	'Unit of Measure.1',	'Unit Price',	
    'Total Committed Value',	'PC Invoice Lock',	'Credit Hold',	'Header Bill Block',	'Item Bill Block',	
    'Bill Plan Bill Block',	'Bill Doc Inv Lock',	'Clarification Case Number',	'External Reference of Billable Item',	
    'Invoice Lock Date',	'Sold-to Id',	'Sold-to Name1',	'Bill-to Id', 'Bill-to Name 1',	'Payer Id',	'Payer Name 1',	
    'Ship-to Id',	'Ship-to Name1',	'Customer PO',	'Country',	'Deal Registration Id',	'Usage',	'ACM Contract',
    'Contract Start Date',	'Contract End Date',	'Contract Created by',	'Invoice Cleared',	'Subprocess',	
    'Created On',	'Contract Account',	'CA Invoice Lock',	'Business Partner',	'Provider Contract',	'Billing Quantity',	
    'Unit of Measure',	'Rate',	'From Date',	'To Date',	'Billable Item Amt',	'Currency',	'Attachment Required',	
    'Print-Relevant',	'Post-Relevant',	'Deferred Revenue Action',	'Rev. Billable Item',	'Reversed Bill Item',	
    'Contract Valid To',	'Bill Line ID',	'Billing Document',	'Billed On',	'Reversed Bill. Doc.', 'Reversal Document',	
    'Revenue Reversed',	'Preceding Document', 'Invoice Order Deleted',	'Invoicing Document',	'Invoicing Category',	
    'Invoiced On'
    ]]

    
    # Create Credit Hold dataframe for items onl credit hold (i.e. no billing blocks)
    CREDIT_HOLD = ZACI_ADUS[ZACI_ADUS['Credit Hold'] == 'Y']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Header Bill Block'] == '']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Item Bill Block'] == '']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Bill Plan Bill Block'] == '']
    CREDIT_HOLD.to_excel(DIR + 'Credit_Hold.xlsx', index=False)

    
    DX = ZACI_ADUS[ZACI_ADUS['Usage'] == '']
    DME = ZACI_ADUS[ZACI_ADUS['Usage'] != '']


    # #print(CREDIT_HOLD.head())
    # print('CREDIT_ZACI_ADUS', CREDIT_HOLD.shape)
    # print('DX', DX.shape)
    # print('DME', DME.shape)
    #ZACI_ADUS.to_excel(DIR + EXCEL, index=False)

    # https://github.com/PyCQA/pylint/issues/3060 pylint: disable=abstract-class-instantiated
    with ExcelWriter(DIR + EXCEL) as writer: #'/ZACI' + 'ZACI_Report.xlsx'
        #write the dataframes to excel
        DX.to_excel(writer, sheet_name='DX', index=False)
        DME.to_excel(writer, sheet_name='DME', index=False)
          
   

##############################################################################
if __name__ == "__main__":
    #zaci_billing_report()
    load_to_pandas()