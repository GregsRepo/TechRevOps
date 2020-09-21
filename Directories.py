###########################################################################

# TechRevOps sharedrive filepath for the automated  reports
FILEPATH = "//du1isi0/order_management/TechRevOps/Reports/Automated/"

# folders in TechRevOps sharedrive for downloaded reports and outputs 
downloads_folder = FILEPATH + "Report_Downloads/"
zaci_folder = FILEPATH + "ZACI/"
provisioning_folder = FILEPATH + "Provisioning/"
output_folder = FILEPATH + "Other"

# zaci and provisioning archive filepaths
prov_archive = provisioning_folder + "archive/"
zaci_archive = zaci_folder + "archive/"

# variables for report downloads file path - "Path + Name"
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
p_status_file = downloads_folder + 'P_Status_Report.txt'

# path for attaching files to e-mail. Also used for archiving
credit_hold_file = zaci_folder + 'Credit_Hold.xlsx'
prov_excel = provisioning_folder + "Provisioning_Report.xlsx"
zaci_excel = zaci_folder + "ZACI_Report.xlsx"


#DESKTOP = "C:/Users/grwillia/OneDrive - Adobe/Desktop/TechRevOps_QE_Reporting/Automated/"

# output to Desktop
# downloads_folder = DESKTOP + "Report_Downloads/"
# zaci_folder = DESKTOP + "ZACI/"
# provisioning_folder = DESKTOP + "Provisioning/"
# output_folder = DESKTOP + "Other"