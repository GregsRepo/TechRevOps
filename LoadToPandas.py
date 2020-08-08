###########################################################################
import numpy as np
import pandas as pd
from pandas import ExcelWriter
import os
import datetime as dt

pd.options.mode.chained_assignment = None  # default='warn'

##############################################################################
# class zaci_dataframe():

#     def __init__(self):
def zaci_dataframe(REGION, FILEPATH):

        '''The SAP Report that this function takes as parameter is split into 2 halfs so that SAP does
        not time out when trying to run the report for a full quarter. In addition to this the 
        downloaded report format is messy because the column data is stacked. In other words the file
        loops back around on itself and we end up with columns stacked on top of one another. Because of
        this a nested for-loop is required. The outside loop sets which file to read in (i.e. which half) and
        the inside loop splits the file into seperate dataframes ([::2] & [1::2]) so as to separate the columns.
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

##############################################################################
def merge_zaci_dataframes(ZACI_ADIR, ZACI_ADUS, ZACI_FOLDER):
    
    '''Change this to take in both dataframe regions and concatenate before performing below operations.
    Will also need to read in DX and DME sheet'''

    # '''Will probably put these directories in their own module and pass them in'''
    # DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/ZACI/'
    # EXCEL = "ZACI_Report.xlsx"

    '''Might need to convert some data points to datetime objects similar to PH dataframes'''
    # Get todays date as a datetime object for use in billing report function
    today = dt.date.today()
    todays_date = today.strftime('%m/%d/%Y') 

    ZACI_COMPLETE = pd.concat([ZACI_ADIR, ZACI_ADUS], ignore_index=True)

    # strip whitespace from all columns
    ZACI_COMPLETE = ZACI_COMPLETE.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # Filter out deleted orders
    ZACI_COMPLETE = ZACI_COMPLETE[ZACI_COMPLETE['Invoice Order Deleted'] != 'X']

    # Add comments field to store order billing block reason
    ZACI_COMPLETE['Comments'] = ''
    # Add comments for line items that will not need review based on billing blocks or a Clarification Case number
    ZACI_COMPLETE['Comments'] = np.where((ZACI_COMPLETE['Clarification Case Number'] != ''), 'Case ' + ZACI_COMPLETE['Clarification Case Number'], 
        np.where((ZACI_COMPLETE['Created On'] == todays_date), 'Created Today',
        np.where((ZACI_COMPLETE['Billed On'] == todays_date), 'Created Today',
        np.where((ZACI_COMPLETE['CA Invoice Lock'] == 'Y'), 'Contr. Acc. Lock', 
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block', 
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('ZM-')), 'ZM - OM Credit Rebill Block',
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block', 
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('ZQ-')), 'ZQ - Provisioning block', 
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block', 
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('13-Final')), 'Final Credit Approval Block',
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('ZH-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('ZS-')), 'Waiting on PO block', 
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('ZS-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('ZS-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('ZV-')), 'Waiting on PO block', 
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('ZV-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('ZV-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('ZW-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Header Bill Block'].str.contains('15-') ), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Item Bill Block'].str.contains('15-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'].str.contains('15-')), 'Waiting on PO block',
        np.where((ZACI_COMPLETE['Header Bill Block'] != ''), 'Review',
        np.where((ZACI_COMPLETE['Item Bill Block'] != ''), 'Review',
        np.where((ZACI_COMPLETE['Bill Plan Bill Block'] != ''), 'Review','Review')))))))))))))))))))))))))))))))


    # Convert Source Transaction ID from object to number for the Join below
    ZACI_COMPLETE['Source Transaction ID'] = pd.to_numeric(ZACI_COMPLETE['Source Transaction ID']) 
    # Drop duplciate Source Transaction ID's. These only occur on Clarification Cases which we don't need anyway            
    ZACI_COMPLETE.drop_duplicates(subset ='Source Transaction ID', keep = False, inplace = True) 
    

    # This section reads in the notes from previous review file using Source Transaction ID as identifier
    DX_NOTES = pd.read_excel(ZACI_FOLDER + "ZACI_Report.xlsx", sheet_name='DX') #'/ZACI_COMPLETE/' + 'ZACI_COMPLETE_Report.xlsx'
    DME_NOTES = pd.read_excel(ZACI_FOLDER + "ZACI_Report.xlsx", sheet_name='DME') 
    DX_NOTES = DX_NOTES[['Notes', 'Source Transaction ID']] # Take only the Notes and Source Transaction Id column
    DME_NOTES = DME_NOTES[['Notes', 'Source Transaction ID']] # Take only the Notes and Source Transaction Id column
    frames = [DX_NOTES, DME_NOTES]
    NOTES = pd.concat(frames) # Concatenate the Notes dataframe together for merge with ZACI_COMPLETE frame
    ZACI_COMPLETE = pd.merge(ZACI_COMPLETE, NOTES, on='Source Transaction ID', how='left') # Merge previous notes with new dataframe based on Source Transaction ID
    

    # Reorder columns
    ZACI_COMPLETE = ZACI_COMPLETE[[
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
    CREDIT_HOLD = ZACI_COMPLETE[ZACI_COMPLETE['Credit Hold'] == 'Y']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Header Bill Block'] == '']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Item Bill Block'] == '']
    CREDIT_HOLD = CREDIT_HOLD[CREDIT_HOLD['Bill Plan Bill Block'] == '']
    CREDIT_HOLD.to_excel(ZACI_FOLDER + 'Credit_Hold.xlsx', index=False)

    # Split complete dataframe by product domain
    DX = ZACI_COMPLETE[ZACI_COMPLETE['Usage'] == '']
    DME = ZACI_COMPLETE[ZACI_COMPLETE['Usage'] != '']

    return DX, DME, CREDIT_HOLD

##############################################################################

def ph_status_dataframe(FILEPATH, PROV_EXCEL, PH_STATUS):

    # Read PH Status data into dataframe
    PH_STATUS_DF = pd.read_csv(PH_STATUS, skiprows=3, sep='|', engine='python')

    # Drop the empty unnamed columns
    cols = [c for c in PH_STATUS_DF.columns if c.lower()[:7] != 'unnamed'] 
    PH_STATUS_DF = PH_STATUS_DF[cols]
    PH_STATUS_DF.drop([0, 0], inplace=True) # Drop first empty rows

     # Rename columns
    PH_STATUS_DF.columns = ['Opportunity ID', 'DR ID', 'ZAV Number.', 'Sold-to', 'End User', 'Re-Seller', 'Deploy To', 'ZAV Create Date', 
                    'Booking date', 'ZAV RFP Date', 'ZAV Provisiong Completion date', 'ZAV Provisiong Error date', 
                    'ZAV User Status', 'PE Error', 'More Messages', 'Order Type', 'Sales Doc.', 'Region', 'Sales Organization', 
                    'Amount', 'Currency', 'Created On', 'Created By', 'Header Block', 'Contract Start Date', 'Contract End Date', 
                    'RFP Date', 'Provisiong Completion date', 'Provisiong Error Date', 'PC Flag', 'User Status']

    # Trim whitespace from all cells
    PH_STATUS_DF = PH_STATUS_DF.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    #PH_STATUS_DF.to_excel(FILEPATH + '/' + 'status.xlsx', index=False)

    # Filter out rejected orders. Used 'ject' becuase we can have mixture of lower/upper case as well as ZAV Rejected
    PH_STATUS_DF = PH_STATUS_DF[~PH_STATUS_DF['Opportunity ID'].str.contains("ject", na=False)]
    PH_STATUS_DF = PH_STATUS_DF[~PH_STATUS_DF['Opportunity ID'].str.contains("JECT", na=False)]

    # Filter out User Status = Canceled
    PH_STATUS_DF = PH_STATUS_DF[~PH_STATUS_DF['User Status'].str.contains("Canceled", na=False)]

    # Create new dataframe at this point for joining with Aging
    JOIN = PH_STATUS_DF[['Opportunity ID', 'ZAV RFP Date', 'ZAV User Status', 'Created On', 'Header Block', 
                            'Contract Start Date', 'Contract End Date']]

    # Change these columns to datetime objects
    today = pd.Timestamp('today').floor('D')
    PH_STATUS_DF['Contract Start Date'] = pd.to_datetime(PH_STATUS_DF['Contract Start Date'])
    PH_STATUS_DF['Created On'] = pd.to_datetime(PH_STATUS_DF['Created On'])

    # Add a Notes column at index 0
    PH_STATUS_DF.insert(loc=0, column='Notes', value='')
    
    # Create seperate dataframes for ZAV User Status equals New, Booking Complete, Prov in Progress, or Prov Error
    NEW = PH_STATUS_DF[(PH_STATUS_DF['ZAV User Status'] == 'New') & (PH_STATUS_DF['Created On'] < today)] 
    BOOKING_COMPLETE = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] == 'Booking Complete']
    PROV_IN_PROGRESS = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] == 'Provisioning in Progress']
    PROVIONING_ERROR = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] == 'Provisioning Error']

    # Add notes to Booking Complete dataframe
    BOOKING_COMPLETE['Notes'] = np.where((BOOKING_COMPLETE['Created On'] == today) , "Created Today",  
                                np.where((BOOKING_COMPLETE['Contract Start Date'] == today) , "Created Today",
                                np.where((BOOKING_COMPLETE['Created On'] > today) , "Future Start Date",  
                                np.where((BOOKING_COMPLETE['Contract Start Date'] > today) , "Future Start Date",
                                np.where((BOOKING_COMPLETE['PC Flag'] == 'Y') , "Credit Block Review",
                                np.where((BOOKING_COMPLETE['Header Block'] == 'ZH : Waiting on PO') , 'On Header Block',
                                np.where((BOOKING_COMPLETE['Header Block'] == 'PP : Provisioning Pending') , 'On Header Block', 
                                'Review')))))))

    # Add notes to PROV_IN_PROGRESS dataframe
    PROV_IN_PROGRESS['Notes'] = np.where((PROV_IN_PROGRESS['Contract Start Date'] > today) , "Future Start Date",
                                np.where((PROV_IN_PROGRESS['Header Block'] == 'ZH : Waiting on PO') , 'On Header Block',
                                np.where((PROV_IN_PROGRESS['Header Block'] == 'PP : Provisioning Pending') , 'On Header Block', 
                                '')))

    
    # Drop Booking Complete items that are also on PO Block
    PH_STATUS_DF.drop(PH_STATUS_DF[(PH_STATUS_DF['ZAV User Status'] == 'Booking Complete')  & (PH_STATUS_DF['Header Block'] != 'ZH : Waiting on PO')].index, inplace=True) 
    
    # Add notes to PH Status dataframe
    PH_STATUS_DF['Notes'] = np.where((PH_STATUS_DF['Contract Start Date'] > today) , "Future Start Date", 'Review')

    # This section reads in the notes from previous review file using Sales Doc as identifier
    STATUS_COMMENTS = pd.read_excel(PROV_EXCEL, sheet_name='Status Review') 
    STATUS_COMMENTS = STATUS_COMMENTS[['Review Comments', 'Sales Doc.']] 
    # Convert Sales Doc. from object to number for the Join below
    STATUS_COMMENTS['Sales Doc.'] = pd.to_numeric(STATUS_COMMENTS['Sales Doc.'])
    PH_STATUS_DF['Sales Doc.'] = pd.to_numeric(PH_STATUS_DF['Sales Doc.']) 
    PH_STATUS_DF = pd.merge(STATUS_COMMENTS, PH_STATUS_DF, on='Sales Doc.', how='right') # Merge previous notes with new dataframe based on Sales Doc
    

    # print('Status New: ', NEW.shape)
    # print('Status PROV_IN_PROGRESS: ', PROV_IN_PROGRESS.shape)
    # print('Status BOOKING_COMPLETE: ', BOOKING_COMPLETE.shape)
    # print('Status PROVIONING_ERROR: ', PROVIONING_ERROR.shape)
    # print('Status PH_STATUS_DF: ', PH_STATUS_DF.shape)

    return BOOKING_COMPLETE, PROV_IN_PROGRESS, NEW, PROVIONING_ERROR, PH_STATUS_DF, JOIN

##############################################################################

def ph_aging_dataframe(FILEPATH, PH_AGING, JOIN):

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

    
    print(PH_AGING_DF.shape)
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
    

    print('Aging New: ', NEW.shape)
    print('Aging Booking Complete: ', BOOKING_COMPLETE.shape)
    print('Aging Provisioning in Progress: ', PROV_IN_PROGRESS.shape)
    print('Aging PROVIONING_ERROR: ', PROVIONING_ERROR.shape)
    print('PH_AGING_DF: ', PH_AGING_DF.shape)

    return BOOKING_COMPLETE, PROV_IN_PROGRESS, NEW, PROVIONING_ERROR, PH_AGING_DF

##############################################################################

def bart_dataframe(FILENAME):

    BART_DF = pd.read_csv(FILENAME, skiprows=4, sep='|', engine='python')
    
    cols = [c for c in BART_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    BART_DF = BART_DF[cols]
    BART_DF.drop([0, 0], inplace=True) # Drop first empty rows

    return BART_DF

##############################################################################
def vfx3_dataframe(FILENAME):

    handle = FILENAME
    if os.stat(handle).st_size == 0:
        d = {'Data': ['No data for input range']}
        VFX3_DF = pd.DataFrame(data=d)
    else:
        VFX3_DF = pd.read_csv(handle, skiprows=3, sep='|', engine='python')
        cols = [c for c in VFX3_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
        VFX3_DF = VFX3_DF[cols]
        VFX3_DF.drop([0, 0], inplace=True) # Drop first empty rows
        

    return VFX3_DF

##############################################################################
def zisexerror_dataframe(FILENAME):
    
    # If empty file create empty dataframe
    handle = FILENAME
    if os.stat(handle).st_size == 0:
        d = {'Data': ['No data for input range']}
        ZISXERROR_DF = pd.DataFrame(data=d)
    else:
        ZISXERROR_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
        cols = [c for c in ZISXERROR_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
        ZISXERROR_DF = ZISXERROR_DF[cols]
        ZISXERROR_DF.drop([0, 0], inplace=True) # Drop first empty rows
    
    return ZISXERROR_DF

##############################################################################

def vuc_dataframe(FILENAME):
    
    # Read V_UC Report into pandas dataframe
    V_UC_DF = pd.read_csv(FILENAME, skiprows=2, sep='\t', engine='python') # 
    
    # Drop the empty unnamed columns
    cols = [c for c in V_UC_DF.columns if c.lower()[:7] != 'unnamed'] 
    V_UC_DF = V_UC_DF[cols]
    
    # Strip whitespace from column names
    V_UC_DF.rename(columns=lambda x: x.strip(), inplace=True)
    
    # Rename Item column to Doc Number
    V_UC_DF.rename(columns={'Item':'Doc Number'}, inplace=True)
    
    # Create new column 'Item Number' to store 
    V_UC_DF['Item No.'] = np.where(V_UC_DF['Doc Number']>=50, '', V_UC_DF['Doc Number'])
    
    # Convert Item No. from object to float
    V_UC_DF['Item No.'] = pd.to_numeric(V_UC_DF['Item No.'], errors='coerce')
    
    # Create new column 'Doc Type' to record whether the document is a ZAV or an Order
    # If > 200M 'ZAV, if < 100M ' Order, else ' ' 
    V_UC_DF['Doc Type'] = np.where((V_UC_DF['Doc Number'] > 200000000) , 'ZAV',   
                 np.where((V_UC_DF['Doc Number'] < 100000000) , '', 
                  np.where((V_UC_DF['Doc Number'] > 100000000) & (V_UC_DF['Doc Number'] < 200000000), 'Order',  
                    '')))   
    
    # Replace item numbers with None in Doc Number column
    V_UC_DF['Doc Number'] = V_UC_DF['Doc Number'].where(V_UC_DF['Doc Number'] > 1000, None) 
    
    # Reorder the columns
    V_UC_DF = V_UC_DF[['Doc Number','Doc Type','Item No.','Short Description','General', 'Delivery', 'BillingDoc', 'Price', \
        'Goods mov.', 'Pck/putaw.', 'Pack']]
    
    return vuc_dataframe

##############################################################################
# if __name__ == "__main__":
    # ph_aging_dataframe("C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports", "PH_Aging_Report.txt", "PH_Status_Report.txt")
#     ph_status_dataframe("C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports", "PH_Status_Report.txt")
