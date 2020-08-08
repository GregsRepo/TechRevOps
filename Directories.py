###########################################################################
# Set the filepath and filename for the downloaded report
FILEPATH = "//du1isi0/order_management/TechRevOps/Reports/Automated/"
#FILEPATH = "C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports"

downloads_folder = FILEPATH + "Report_Downloads/"
zaci_folder = FILEPATH + "ZACI/"
provisioning_folder = FILEPATH + "Provisioning/"
output_folder = FILEPATH + "Other"

#zaci_excel = zaci_folder + "ZACI_Report.xlsx"
prov_excel = provisioning_folder + "Provisioning_Report.xlsx"



ph_aging_file = downloads_folder + "PH_Aging_Report.txt"
ph_status_file = downloads_folder + "PH_Status_Report.txt"
bart_error_file = downloads_folder + "BART_Error_Report.txt"
bart_duplicate_file = downloads_folder + 'BART_Duplicate_Report.txt'
bart_no_provision_file = downloads_folder + 'BART_No_Provisioning.txt'
vfx3_ihc_file = downloads_folder + 'VFX3_IHC_Report.txt'
vfx3_adir_file = downloads_folder + 'VFX3_ADIR_Report.txt'
vfx3_adus_file = downloads_folder + 'VFX3_ADUS_Report.txt'
v_uc_file = downloads_folder + 'V_UC_Report.txt'
zisxerror_file = downloads_folder + 'ZISXERROR_Report.txt'