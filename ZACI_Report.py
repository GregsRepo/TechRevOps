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
class ZACI():

    def __init__(self, REGION, FILEPATH):
        self.REGION = REGION
        self.FILEPATH = FILEPATH
        
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

##############################################################################
def load_to_pandas(REGION):

        '''The SAP Report that this class takes as parameter is split into 2 halfs so that SAP does
        not time out when trying to run the report for a full quarter. In addition to this the 
        downlaoded report format is messy because the column data is stacked. In other words the file
        loops back around on itself and we end up with columns stacked on top of one another. Because of
        this a nested for loop is required. The outside loop sets which file to read in (i.e. which half) and
        the inside loop splits the file into seperate dataframes ([::2] & [1::2]) so as separate the columns.
        It then concatenates the seperate dataframes into one legible report. It will then do the same for the 
        next file (i.e. 2nd half) and then concatenates both dataframes into one full quarter report.'''
    
        # Declare empty dataframes. 1 to concatenate the half reports together 
        # and the other to concatenate both halfs into one full report
        HALF = pd.DataFrame()
        ZACI = pd.DataFrame()
        for i in range(2):
            # Set which half of the SAP report to read in
            if i == 0:
                FILENAME = 'ZACI_Report_' + REGION + '_1st_half.txt'
            else:
                FILENAME = 'ZACI_Report_' + REGION + '_2nd_half.txt'
            for j in range(2):
                if j == 0:
                    ZACI_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
                    # drop the empty 'unnamed' columns
                    cols = [c for c in ZACI_DF.columns if c.lower()[:7] != 'unnamed'] 
                    ZACI_DF = ZACI_DF[cols]
                    # drop the the first 2 rows from ZACI_DF and retrieve every 2nd row and then reset index
                    ZACI_DF = ZACI_DF.drop([0, 1])
                    ZACI_DF = ZACI_DF.iloc[::2]
                    ZACI_DF = ZACI_DF.reset_index(drop=True)
                else:
                    ZACI_DF1 = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
                    # drop the the first 2 rows from ZACI_DF and retrieve every 2nd row and then reset index
                    ZACI_DF1 = ZACI_DF1.drop([1, 2])
                    ZACI_DF1.columns = ZACI_DF1.iloc[0]
                    ZACI_DF1 = ZACI_DF1.iloc[1::2]
                    ZACI_DF1 = ZACI_DF1.reset_index(drop=True)
                    # Concatenate the the 2 dataframes to give us one legible half of the report
                    HALF = pd.concat([ZACI_DF, ZACI_DF1], axis=1, sort=False)
            # Drop Nan columns and rows and strip whitespace from dataframes before concatenating (otherwise we get misaligned shape error)
            HALF.dropna(axis = 1, how ='all', inplace = True) 
            HALF.dropna(axis = 0, how ='all', inplace = True) 
            HALF.columns = HALF.columns.str.strip()
            # Concatenate the reports together. The first loop will concatenate the first half to an empty dataframe. 
            # The second loop will concatenate the second half to the first half
            ZACI = pd.concat([ZACI, HALF], sort=False)
        
    
        return ZACI    
            
        #filter_dataframe(ZACI)

##############################################################################
def filter_dataframe(ADIR, ADUS):

    ZACI = pd.concat([ADIR, ADUS], sort=False)

    '''Change this to take in both dataframe REGIONs and concatenate before performing below operations.
    Will also need to read in DX and DME sheet'''

    # strip whitespace from all columns
    ZACI = ZACI.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # Filter out deleted orders
    ZACI = ZACI[ZACI['Invoice Order Deleted'] != 'X']

    # Add comments field to store order billing block reason
    ZACI['Comments'] = ''
    # Add comments for line items that will not need review based on billing blocks or a Clarification Case number
    ZACI['Comments'] = np.where((ZACI['Clarification Case Number'] != ''), 'Case ' + ZACI['Clarification Case Number'], 
        np.where((ZACI['Created On'] == todays_date), 'Created Today',
        np.where((ZACI['Billed On'] == todays_date), 'Created Today',
        np.where((ZACI['CA Invoice Lock'] == 'Y'), 'Contr. Acc. Lock', 
        np.where((ZACI['Header Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block', 
        np.where((ZACI['Item Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block',
        np.where((ZACI['Header Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block', 
        np.where((ZACI['Item Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block', 
        np.where((ZACI['Header Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block', 
        np.where((ZACI['Item Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block',
        np.where((ZACI['Header Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI['Item Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI['Header Bill Block'].str.contains('ZS-')), 'Waiting on PO block', 
        np.where((ZACI['Item Bill Block'].str.contains('ZS-')), 'Waiting on PO block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('ZS-')), 'Waiting on PO block',
        np.where((ZACI['Header Bill Block'].str.contains('ZV-')), 'Waiting on PO block', 
        np.where((ZACI['Item Bill Block'].str.contains('ZV-')), 'Waiting on PO block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('ZV-')), 'Waiting on PO block',
        np.where((ZACI['Header Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI['Item Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI['Header Bill Block'].str.contains('15-') ), 'Waiting on PO block',
        np.where((ZACI['Item Bill Block'].str.contains('15-')), 'Waiting on PO block',
        np.where((ZACI['Bill Plan Bill Block'].str.contains('15-')), 'Waiting on PO block',
        np.where((ZACI['Header Bill Block'] != ''), 'Review',
        np.where((ZACI['Item Bill Block'] != ''), 'Review',
        np.where((ZACI['Bill Plan Bill Block'] != ''), 'Review','Review')))))))))))))))))))))))))))))))


    # Convert Source Transaction ID from object to number for the Join below
    ZACI['Source Transaction ID'] = pd.to_numeric(ZACI['Source Transaction ID']) 
    # Drop duplciate Source Transaction ID's. These only occur on Clarification Cases which we don't need anyway            
    ZACI.drop_duplicates(subset ='Source Transaction ID', keep = False, inplace = True) 
    print('ZACI Report', ZACI.shape)

    # This section reads in the notes from previous review file using Source Transaction ID as identifier
    DX_NOTES = pd.read_excel(DIR + EXCEL, sheet_name='DX') #'/ZACI/' + 'ZACI_Report.xlsx'
    DME_NOTES = pd.read_excel(DIR + EXCEL, sheet_name='DME') 
    DX_NOTES = DX_NOTES[['Notes', 'Source Transaction ID']] # Take only the Notes and Source Transaction Id column
    DME_NOTES = DME_NOTES[['Notes', 'Source Transaction ID']] # Take only the Notes and Source Transaction Id column
    frames = [DX_NOTES, DME_NOTES]
    NOTES = pd.concat(frames) # Concatenate the Notes dataframe together for merge with ZACI frame
    ZACI = pd.merge(ZACI, NOTES, on='Source Transaction ID', how='left') # Merge previous notes with new dataframe based on Source Transaction ID
    

    # Reorder columns
    ZACI = ZACI[[
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

    
    # Create Credit Hold dataframe for items on credit hold (i.e. no billing blocks)
    CREDIT_HOLD = ZACI[ZACI['Credit Hold'] == 'Y']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Header Bill Block'] == '']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Item Bill Block'] == '']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Bill Plan Bill Block'] == '']
    CREDIT_HOLD.to_excel(DIR + 'Credit_Hold.xlsx', index=False)

    # Create seperate dtaframes for DX and DME
    DX = ZACI[ZACI['Usage'] == '']
    DME = ZACI[ZACI['Usage'] != '']

    CREDIT_HOLD.to_excel(DIR + 'Credit_Hold.xlsx', index=False)
    DX.to_excel(DIR + EXCEL, sheet_name='DX', index=False)
    DME.to_excel(DIR + EXCEL, sheet_name='DME', index=False)

    # # https://github.com/PyCQA/pylint/issues/3060 pylint: disable=abstract-class-instantiated
    # with ExcelWriter(DIR + EXCEL) as writer: 
    #     #write the dataframes to excel
    #     DX.to_excel(writer, sheet_name='DX', index=False)
    #     DME.to_excel(writer, sheet_name='DME', index=False)
          
   

##############################################################################
# if __name__ == "__main__":
#     #zaci_billing_report()
#     load_to_pandas()