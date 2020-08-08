import SAP_Reports as sap
import Directories as Dir 
import LoadToPandas as ltp
import WriteToFile as wtf
import datetime as dt

###########################################################################
# Get todays date as a datetime object for use in billing report function
today = dt.date.today()


###########################################################################


# ADUS = SAP.ZACI('ADUS', Dir.downloads_folder)
# ADIR = SAP.ZACI('ADIR', Dir.downloads_folder)
# ADUS = SAP.load_to_pandas('ADUS')
# print(ADUS.shape)


if __name__ == "__main__":
    
    # sap.zaci('ADUS', Dir.downloads_folder)
    # sap.zaci('ADIR', Dir.downloads_folder)
    # ph_aging_file = sap.ph_aging(Dir.downloads_folder)
    # ph_status_file = sap.ph_status(Dir.downloads_folder)
    # bart_error_file = sap.bart('1', Dir.downloads_folder)
    # bart_duplicate_file = sap.bart('3', Dir.downloads_folder)
    # bart_no_provision_file = sap.bart('7', Dir.downloads_folder)
    # vfx3_ihc_file = sap.vfx3('I001', Dir.downloads_folder)
    # vfx3_adir_file = sap.vfx3('D001', Dir.downloads_folder)
    # vfx3_adus_file = sap.vfx3('0001', Dir.downloads_folder)
    # v_uc_file = sap.v_uc(Dir.downloads_folder)
    # zisxerror_file = sap.zisxerror(Dir.downloads_folder)

    
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
    
    '''Might chnage these to look at downlaod folder for relevant files rather than use the variables above'''
    # zaci_adus_df = ltp.zaci_dataframe('ADUS', Dir.downloads_folder)
    # zaci_adir_df = ltp.zaci_dataframe('ADIR', Dir.downloads_folder)
    # dx, dme, credit_hold = ltp.merge_zaci_dataframes(zaci_adir_df, zaci_adus_df, Dir.zaci_folder)
    ph_status_dataframes = ltp.ph_status_dataframe(Dir.downloads_folder, Dir.prov_excel, Dir.ph_status_file)
    JOIN = ph_status_dataframes[5]
    ph_aging_dataframes = ltp.ph_aging_dataframe(Dir.downloads_folder, Dir.ph_aging_file, JOIN)
    # bart_error_df = ltp.bart_dataframe(Dir.bart_error_file)
    # bart_duplicate_df = ltp.bart_dataframe(Dir.bart_duplicate_file)
    # bart_no_provision_df = ltp.bart_dataframe(Dir.bart_no_provision_file)
    # vfx3_ihc_df = ltp.vfx3_dataframe( Dir.vfx3_ihc_file)
    # vfx3_adir_df = ltp.vfx3_dataframe(Dir.vfx3_adir_file)
    # vfx3_adus_df = ltp.vfx3_dataframe(Dir.vfx3_adus_file)
    # v_uc_df = ltp.vfx3_dataframe(Dir.v_uc_file) 
    # zisxerror_df = ltp.zisexerror_dataframe(Dir.zisxerror_file)

    
    # wtf.zaci_to_excel(dx, dme, credit_hold, Dir.zaci_folder)
    wtf.ph_to_excel(ph_status_dataframes, ph_aging_dataframes, Dir.provisioning_folder)
    # wtf.excel(bart_error_df, Dir.output_folder, Dir.bart_error_file[72:-4] + '.xlsx')
    # wtf.excel(bart_duplicate_df, Dir.output_folder, Dir.bart_duplicate_file[72:-4] + '.xlsx')
    # wtf.excel(bart_no_provision_df, Dir.output_folder, Dir.bart_no_provision_file[72:-4] + '.xlsx')
    # wtf.excel(vfx3_ihc_df, Dir.output_folder, Dir.vfx3_ihc_file[72:-4] + '.xlsx')
    # wtf.excel(vfx3_adir_df, Dir.output_folder, Dir.vfx3_adir_file[72:-4] + '.xlsx')
    # wtf.excel(vfx3_adus_df, Dir.output_folder, Dir.vfx3_adus_file[72:-4] + '.xlsx')
    # wtf.excel(v_uc_df, Dir.output_folder, Dir.v_uc_file[72:-4] + '.xlsx') 
    # wtf.excel(zisxerror_df, Dir.output_folder, Dir.zisxerror_file[72:-4] + '.xlsx')
