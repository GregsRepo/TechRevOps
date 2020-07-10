###########################################################################
import numpy as np
import pandas as pd
from pandas import ExcelWriter
import os
import datetime as dt

##############################################################################
# class zaci_dataframe():

#     def __init__(self):
def zaci_dataframe(REGION, FILEPATH):

        '''The SAP Report that this class takes as parameter is split into 2 halfs so that SAP does
        not time out when trying to run the report for a full quarter. In addition to this the 
        downloaded report format is messy because the column data is stacked. In other words the file
        loops back around on itself and we end up with columns stacked on top of one another. Because of
        this a nested for loop is required. The outside loop sets which file to read in (i.e. which half) and
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

def ph_status_dataframe(FILEPATH, PH_STATUS):

    PH_STATUS_DF = pd.read_csv(FILEPATH + '/' + PH_STATUS, skiprows=3, sep='|', engine='python')
    
    print('PH_STATUS_DF: ', PH_STATUS_DF.shape)

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

    PH_STATUS_DF.to_excel(FILEPATH + '/' + 'status.xlsx', index=False)

    # Filter out rejected orders
    PH_STATUS_DF = PH_STATUS_DF.drop(PH_STATUS_DF[(PH_STATUS_DF['Opportunity ID'] != 'Rejected') 
                    & (PH_STATUS_DF['Opportunity ID'] != 'ZAV Reject') 
                    & (PH_STATUS_DF['Opportunity ID'] != 'Reject')
                    & (PH_STATUS_DF['Opportunity ID'] != 'Canceled')].index)

    # Change these columns to datetime objects
    #date = dt.date.today()
    today = pd.Timestamp('today').floor('D')
    PH_STATUS_DF['Contract Start Date'] = pd.to_datetime(PH_STATUS_DF['Contract Start Date'])
    PH_STATUS_DF['Created On'] = pd.to_datetime(PH_STATUS_DF['Created On'])

    # Add a Notes column at index 0
    PH_STATUS_DF.insert(loc=0, column='Notes', value='')
    
    NEW = PH_STATUS_DF[(PH_STATUS_DF['ZAV User Status'] == 'New') & (PH_STATUS_DF['Created On'] < today)] 
    #PH_STATUS_DF = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] != 'New']

    BOOKING_COMPLETE = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] == 'Booking Complete']
    # PH_STATUS_DF = PH_STATUS_DF[(PH_STATUS_DF['Header Block'] != 'ZH : Waiting on PO') 
    #                             & (PH_STATUS_DF['Contract Start Date'] <= today)]


    PROV_IN_PROGRESS = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] == 'Provisioning in Progress']
    #PH_STATUS_DF = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] != 'Provisioning in Progress']

    PROVIONING_ERROR = PH_STATUS_DF[PH_STATUS_DF['ZAV User Status'] == 'Provisioning Error']

    print('PH_STATUS_DF: ', PH_STATUS_DF.shape)
    '''Everything gets dropped here'''
    # Filter PH Status Dataframe based on conditions
    PH_STATUS_DF = PH_STATUS_DF[(PH_STATUS_DF['ZAV User Status'] == 'Booking Complete') 
                                & (PH_STATUS_DF['Header Block'] != 'ZH : Waiting on PO') 
                                & (PH_STATUS_DF['Contract Start Date'] <= today)]

    
    # Add notes to Booking Complete dataframe
    PH_STATUS_DF['Notes'] = np.where((PH_STATUS_DF['Contract Start Date'] > today) , "Future Start Date", 'Review')

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



    # print('New: ', NEW.shape)
    # print('PROV_IN_PROGRESS: ', PROV_IN_PROGRESS.shape)
    # print('BOOKING_COMPLETE: ', BOOKING_COMPLETE.shape)
    # print('PROVIONING_ERROR: ', PROVIONING_ERROR.shape)
    # print('PH_STATUS_DF: ', PH_STATUS_DF.shape)

    return NEW, BOOKING_COMPLETE, PROV_IN_PROGRESS, PROVIONING_ERROR, PH_STATUS_DF

##############################################################################

def ph_aging_dataframe(FILEPATH, PH_AGING, PH_STATUS_DF):


    '''This needs to be rewritten to read in both provisining dataframes and match the Alteryx report
    Might create a seperate function for PH_status. Only reading in here so as to get the contract start date.
    Will need to create PH_Status dataframe first and then pass that into aging function'''

    PH_AGING_DF = pd.read_csv(FILEPATH + '/' + PH_AGING, skiprows=3, sep='|', engine='python')
    
    print('PH_AGING_DF: ', PH_AGING_DF.shape)

    # Drop the empty unnamed columns
    cols = [c for c in PH_STATUS_DF.columns if c.lower()[:7] != 'unnamed'] 
    PH_STATUS_DF = PH_STATUS_DF[cols]
    if PH_STATUS_DF.empty:
        pass 
    else:
        PH_STATUS_DF.drop([0, 0], inplace=True) # Drop first empty rows

        
    # Take columns from Ph Status for joining with PH Aging
    JOIN = PH_STATUS_DF[['Opportunity ID', 'ZAV RFP Date', 'ZAV User Status', 'Created On', 'Header Block', 
                            'Contract Start Date', 'Contract End Date']]
    SALES_DOC = PH_STATUS_DF['Sales Doc.']
    
    # Drop the empty unnamed columns from PH Aging dataframe
    cols = [c for c in PH_AGING_DF.columns if c.lower()[:7] != 'unnamed'] 
    PH_AGING_DF = PH_AGING_DF[cols]
    PH_AGING_DF.drop([0, 0], inplace=True) # Drop first empty rows

    # Rename columns
    PH_AGING_DF.columns = ['Opportunity ID', 'Region', 'DR Number', 'Sales Org', 'EU Country Code', 'Currency', 'Customer', 'End User',
                            'EU Cust Name', 'Sales Order Created By', 'Document No.', 'Sales Doc.', 'Doc Type', 'Amount', 'After PC', 
                            'New', 'Booking Complete', 'Provisioning in Progress', 'Provisioning Completed', 'Provisioning Error',
                            'Total No. of Days', 'Create Date(ZCC)', 'Create Date(ZAV)', 'Last Status Date']
    
    # Add a Notes column at index 0
    PH_AGING_DF.insert(loc=0, column='Notes', value='')

    # Join data taken from PH Status Report to the Aging Report
    PH_AGING_DF = pd.concat([PH_AGING_DF, JOIN], sort=False)

    # Trim whitespace from all cells
    PH_AGING_DF = PH_AGING_DF.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    PH_AGING_DF.to_excel(FILEPATH + '/' + 'aging.xlsx', index=False)

    # Filter out rejected orders
    PH_AGING_DF = PH_AGING_DF.drop(PH_AGING_DF[(PH_AGING_DF['Opportunity ID'] != 'Rejected') 
                    & (PH_AGING_DF['Opportunity ID'] != 'ZAV Reject') 
                    & (PH_AGING_DF['Opportunity ID'] != 'Reject')].index)


    # Change these columns to datetime objects
    #date = dt.date.today()
    today = pd.Timestamp('today').floor('D')
    PH_AGING_DF['Create Date(ZCC)'] = pd.to_datetime(PH_AGING_DF['Create Date(ZCC)'])
    PH_AGING_DF['Contract Start Date'] = pd.to_datetime(PH_AGING_DF['Contract Start Date'])
    PH_AGING_DF['Create Date(ZAV)'] = pd.to_datetime(PH_AGING_DF['Create Date(ZAV)'])
    
    NEW = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] == 'New']
    #PH_AGING_DF = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] != 'New']

    BOOKING_COMPLETE = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] == 'Booking Complete']
    #PH_AGING_DF = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] != 'Booking Complete']

    PROV_IN_PROGRESS = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] == 'Provisioning in Progress']
    #PH_AGING_DF = PH_AGING_DF[PH_AGING_DF['ZAV User Status'] != 'Provisioning in Progress']

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
    
    print('PH_AGING_DF: ', PH_AGING_DF.shape)

    '''Everything gets dropped here'''
    # Filter PH Aging Dataframe based on conditions
    PH_AGING_DF = PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'New') 
                                & (PH_AGING_DF['Create Date(ZCC)'].dt.date < today) ]
    

    PH_AGING_DF = PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'Booking Complete') 
                                & (PH_AGING_DF['Header Block'] != 'ZH : Waiting on PO') 
                                | (PROV_IN_PROGRESS['Header Block'] == 'PP : Provisioning Pending')]

    PH_AGING_DF = PH_AGING_DF[(PH_AGING_DF['ZAV User Status'] == 'Provisioning in Progress') 
                                & (PH_AGING_DF['Create Date(ZCC)'].dt.date < today) 
                                & (PH_AGING_DF['Contract Start Date'] < today)]

    
    # Drop anything on these header bill block
    PH_AGING_DF = PH_AGING_DF[(PH_AGING_DF['Header Block'] != 'ZH : Waiting on PO') | (PH_AGING_DF['Header Block'] != 'PP : Provisioning Pending') ]
    
    # Find out why we do this?
    PH_AGING_DF = pd.merge(SALES_DOC, PH_AGING_DF, on='Sales Doc.')

    # Add a note for anything that is Waiting on PO or has a Future Start Date
    PH_AGING_DF['Notes'] = np.where((PH_AGING_DF['Header Block'] == "ZH : Waiting on PO") , "Billing Block",   
                 np.where((PH_AGING_DF['Contract Start Date'].dt.date > today) , 'Future Start Date',
                  'Review'))   

    # print('New: ', NEW.shape)
    # print('PROV_IN_PROGRESS: ', PROV_IN_PROGRESS.shape)
    # print('BOOKING_COMPLETE: ', BOOKING_COMPLETE.shape)
    # print('PROVIONING_ERROR: ', PROVIONING_ERROR.shape)

    return NEW, BOOKING_COMPLETE, PROV_IN_PROGRESS, PROVIONING_ERROR, PH_AGING_DF

##############################################################################

def bart_dataframe(FILEPATH, FILENAME):

    BART_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=4, sep='|', engine='python')
    
    cols = [c for c in BART_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    BART_DF = BART_DF[cols]
    BART_DF.drop([0, 0], inplace=True) # Drop first empty rows

    return BART_DF

##############################################################################
def vfx3_dataframe(FILEPATH, FILENAME):

    handle = FILEPATH + '/' + FILENAME
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
def zisexerror_pandas(FILEPATH, FILENAME):
    
    ZISXERROR_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=3, sep='|', engine='python')
    
    cols = [c for c in ZISXERROR_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
    ZISXERROR_DF = ZISXERROR_DF[cols]
    ZISXERROR_DF.drop([0, 0], inplace=True) # Drop first empty rows
    
    return ZISXERROR_DF

##############################################################################

def vuc_dataframe(FILEPATH, FILENAME):
    
    # Read V_UC Report into pandas dataframe
    V_UC_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=2, sep='\t', engine='python') # 
    
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
#     # ph_aging_dataframe("C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports", "PH_Aging_Report.txt", "PH_Status_Report.txt")
#     ph_status_dataframe("C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports", "PH_Status_Report.txt")
