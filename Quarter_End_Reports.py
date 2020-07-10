#import SAP
import SAP_Reports as sap
import LoadToPandas as ltp
import WriteToFile as wtf
import datetime as dt


# Get todays date as a datetime object for use in billing report function
today = dt.date.today()

###########################################################################
# Set the filepath and filename for the downloaded report
#FILEPATH = "//du1isi0/order_management/TechRevOps/Reports/SAP_Reports/Report Downloads"
FILEPATH = "C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/ZACI/'

###########################################################################



# ADUS = SAP.ZACI('ADUS', FILEPATH)
# ADIR = SAP.ZACI('ADIR', FILEPATH)
# ADUS = SAP.load_to_pandas('ADUS')
# print(ADUS.shape)


if __name__ == "__main__":
    
    sap.zaci('ADUS', FILEPATH)
    sap.zaci('ADIR', FILEPATH)
    # ph_aging_file = sap.ph_aging(FILEPATH)
    # ph_status_file = sap.ph_status(FILEPATH)
    # bart_error_file = sap.bart('1', FILEPATH)
    # bart_duplicate_file = sap.bart('3', FILEPATH)
    # bart_no_provision_file = sap.bart('7', FILEPATH)
    # vfx3_ihc_file = sap.vfx3('I001', FILEPATH)
    # vfx3_adir_file = sap.vfx3('D001', FILEPATH)
    # vfx3_adus_file = sap.vfx3('0001', FILEPATH)
    # v_uc_file = sap.v_uc(FILEPATH)
    # zisxerror_file = sap.zisxerror(FILEPATH)

    
    # print(ph_aging_file)
    # print(ph_status_file)
    # print(bart_error_file)
    # print(bart_duplicate_file)
    # print(bart_no_provision_file)
    # print(vfx3_ihc_file)
    # print(vfx3_adir_file)
    # print(vfx3_adus_file)
    # print(v_uc_file)
    # print(zisxerror_file)
    
    # zaci_adus_df = ltp.zaci_dataframe('ADUS', FILEPATH)
    # zaci_adir_df = ltp.zaci_dataframe('ADIR', FILEPATH)
    # ph_status_file = "PH_Status_Report.txt"
    # ph_aging_file = "PH_Aging_Report.txt"
    # ph_status_dataframes = ltp.ph_status_dataframe(FILEPATH, ph_status_file)
    # ph_status_df = ph_status_dataframes[4]
    # ph_aging_dataframes = ltp.ph_aging_dataframe(FILEPATH, ph_aging_file, ph_status_df)
    # bart_error_df = ltp.bart_dataframe(FILEPATH, bart_error_file)
    # bart_duplicate_df = ltp.bart_dataframe(FILEPATH, bart_duplicate_file)
    # bart_no_provision_df = ltp.bart_dataframe(FILEPATH, bart_no_provision_file)
    # vfx3_ihc_df = ltp.vfx3_dataframe(FILEPATH, vfx3_ihc_file)
    # vfx3_adir_df = ltp.vfx3_dataframe(FILEPATH, vfx3_adir_file)
    # vfx3_adus_df = ltp.vfx3_dataframe(FILEPATH, vfx3_adus_file)
    # v_uc_df = ltp.vfx3_dataframe(FILEPATH, v_uc_file) 
    # zisxerror_df = ltp.zisexerror_pandas(FILEPATH, zisxerror_file)

    
    # wtf.excel(zaci_adus_df, DIR, 'ZACI_Report.xlsx')
    # wtf.excel(ph_aging_df, DIR, ph_aging_file[:-4] + '.xlsx')
    # wtf.excel(ph_status_df, DIR, ph_status_file[:-4] + '.xlsx')
    # wtf.ph_to_excel(ph_status_dataframes, ph_aging_dataframes, DIR)
    # wtf.excel(bart_error_df, DIR, bart_error_file[:-4] + '.xlsx')
    # wtf.excel(bart_duplicate_df, DIR, bart_duplicate_file[:-4] + '.xlsx')
    # wtf.excel(bart_no_provision_df, DIR, bart_no_provision_file[:-4] + '.xlsx')
    # wtf.excel(vfx3_ihc_df, DIR, vfx3_ihc_file[:-4] + '.xlsx')
    # wtf.excel(vfx3_adir_df, DIR, vfx3_adir_file[:-4] + '.xlsx')
    # wtf.excel(vfx3_adus_df, DIR, vfx3_adus_file[:-4] + '.xlsx')
    # wtf.excel(v_uc_df, DIR, v_uc_file[:-4] + '.xlsx') 
    # wtf.excel(zisxerror_df, DIR, zisxerror_file[:-4] + '.xlsx')
