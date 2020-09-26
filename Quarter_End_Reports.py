import functools, random, smtplib, string, sys, datetime, os
from shutil import copyfile
import SAP_Reports as sap
import Directories as Dir 
import LoadToPandas as ltp
import WriteToFile as wtf
import Dataframe_Length
import EmailMessage
import QEndEmail, HtmlEmail
import datetime as dt


'''TO DO: Look at how to handle when the Provisioning Report "Status Review" tab does not write out to excel. 
When this happens it fails to write out the ph aging data and the next time the script is run it fails completely
because there is no status review tab to read previous comments from. Not a major concern as Provisioning Aging
data is not soemthing we are too concerned with and the file can be corrected before running the script again.

For improvements, you want to look at sending an e-mail seperately for the credit file and attaching files to 
both the quarter end report and credit e-mail

Might also look at general excpetion handling as bugs are discovered'''

###############################################################################################################################
def LogIn():
    # Global value to take the users ldap at login and pass it to the email fucntion at end of this script
    global ldap
    ldap = input('Please enter your LDAP or enter 0 to  ').lower()
    ListOfAuthUsers = ['grwillia','cinnide','tporter','tracyl','dpurcell', 'gholbroo']
    if ldap == '0':
        print('Exit program')
        exit()
    elif(ldap in ListOfAuthUsers):
        token =  (''.join(random.choice(string.ascii_uppercase + string.digits) for x in range(10)))
        SendEmail = ldap + '@Adobe.com'
        FromEmail = 'DoNotReply@Adobe.com'
        server = smtplib.SMTP('namail.corp.adobe.com', 25)
        server.sendmail(FromEmail, SendEmail, token)
        print('Token Sent . . .')
        inpt = input('Hello,\nPlease check your email and enter your token (Case Sensitive) or press 0 to exit: ')
        if (token == inpt):
            print('Token Accepted!')
            # Call Menu
            get_menu()
        elif(inpt == '0'):
            print('Exit Utility.')
            exit()
        else:
            print('Incorrect Token. Please try again or press 0 to exit' )
            LogIn()
    else:
        print('Access Denied. Please contact Grp-TechRevOps@adobe.com for privilages.')
        LogIn()

###############################################################################################################################
def get_menu():
    

    option = ''
    while option != "0":
        option = input('''\nChoose an option(Pick a number. Or '0' to exit  
        0: Exit
        1: Full Quarter End Report (includes download of new SAP Reports)
        2: Reload output files (just reload files from previous SAP downloads)
        3: Run ZACI Reports (only run the ZACI report)
        4: Run Provisioning Reports (only run the Provisioning Reports)
        Option: ''')

        if option == "0":
            sys.exit()
        elif option == "1":
            run_sap_reports()
        elif option == "2":
            call_load_to_pandas()
        elif option == "3":
            run_zaci_reports()
        elif option == "4":
            run_provisioning_reports()
            

        menu = ["1", "2", "3", "4"]
        if option not in menu:
            print('\nPick a number between 1 & 4: ')

###############################################################################################################################
def run_sap_reports():

    # Download the various Reports by calling the SAP_Reports function
    print('\nSAP Reports:')
    sap.zaci('ADUS', Dir.downloads_folder)
    sap.zaci('ADIR', Dir.downloads_folder)
    sap.ph_aging(Dir.downloads_folder)
    sap.ph_status(Dir.downloads_folder)
    sap.bart('1', Dir.downloads_folder)
    sap.bart('3', Dir.downloads_folder)
    sap.bart('7', Dir.downloads_folder)
    sap.vfx3('I001', Dir.downloads_folder)
    sap.vfx3('D001', Dir.downloads_folder)
    sap.vfx3('0001', Dir.downloads_folder)
    sap.v_uc(Dir.downloads_folder)
    sap.zisxerror(Dir.downloads_folder)
    sap.p_status(Dir.downloads_folder)

    call_load_to_pandas()

###############################################################################################################################
def archive_reports(src, archive):

    # Datetime formatting for apending to archived file
    now = datetime.datetime.now().time()
    time = now.strftime("%f")

    # set source path and take filename from path
    #src = path
    file_name = os.path.basename(src)

    try:
        # copy file to archive folder
        print('Archiving ' + file_name)
        archive_file = file_name[:-5] + '_' + time + '.xlsx'
        dst = archive + archive_file
        copyfile(src, dst)
    except:
        print('Error encountered archiving ' + file_name)
        pass

###############################################################################################################################
def run_zaci_reports():
    
    print('\nRunning ZACI SAP Reports:')
    sap.zaci('ADUS', Dir.downloads_folder)
    sap.zaci('ADIR', Dir.downloads_folder)

    # Create dataframes for the ZACI Reports by calling the LoadToPandas function 
    print('Loading ZACI report...')
    try:
        zaci_adus_df = ltp.zaci_dataframe('ADUS', Dir.downloads_folder)
    except Exception as e:
        input('Error encountered loading ZACI ADUS dataframe.\n' + str(e))
        exit()
    
    try:
        zaci_adir_df = ltp.zaci_dataframe('ADIR', Dir.downloads_folder)
    except Exception as e:
        input('Error encountered loading ZACI ADIR dataframe.\n' + str(e))
        exit()
    
    try:
        dx, dme, credit_hold = ltp.merge_zaci_dataframes(zaci_adir_df, zaci_adus_df, Dir.zaci_folder)
    except Exception as e:
        input('Error encountered merging ZACI dataframes.\n' + str(e))
        exit()

    # archive previously ran zaci files
    archive_reports(Dir.zaci_excel, Dir.zaci_archive)
    archive_reports(Dir.credit_hold_file, Dir.zaci_archive)

    # Write the various dataframes to excel by calling the WriteToFile function
    print('Writing ZACI report out to Excel...')
    wtf.zaci_to_excel(dx, dme, credit_hold, Dir.zaci_folder)

###############################################################################################################################
def run_provisioning_reports():

    # Download the Provisioning Reports by calling the SAP_Reports function
    sap.ph_aging(Dir.downloads_folder)
    sap.ph_status(Dir.downloads_folder)

    # Create dataframes for the Provisioning Reports by calling the LoadToPandas function  
    # PH Status Report
    print('Loading Provisioning Status report...')
    try:
        ph_status_returns = ltp.ph_status_dataframe(Dir.prov_excel, Dir.ph_status_file)
        ph_status_dataframes = ph_status_returns[:5]
        JOIN = ph_status_returns[5]
    except Exception as e:
        input('Error encountered loading PH Status dataframe.\n' + str(e))
        exit()
    
    # PH Aging Report
    try:
        print('Loading Provisioning Aging report...')
        ph_aging_dataframes = ltp.ph_aging_dataframe(Dir.prov_excel, Dir.ph_aging_file, JOIN)
    except Exception as e:
        input('Error encountered loading PH Aging dataframe.\n' + str(e))
        exit()
    
    # archive previously ran provisioning file
    archive_reports(Dir.prov_excel, Dir.prov_archive)

    # Write the various dataframes to excel by calling the WriteToFile function
    print('Writing Provisioning Report to excel...')
    wtf.ph_to_excel(ph_status_dataframes, ph_aging_dataframes, Dir.provisioning_folder)

###########################################################################
def call_load_to_pandas():

    print('\nWriting Reports to memory...:')
    # Create dataframes for the ZACI Reports by calling the LoadToPandas function 
    print('Loading ZACI report...')
    try:
        zaci_adus_df = ltp.zaci_dataframe('ADUS', Dir.downloads_folder)
    except Exception as e:
        input('Error encountered loading ZACI ADUS dataframe.\n' + str(e))
        exit()
    try:
        zaci_adir_df = ltp.zaci_dataframe('ADIR', Dir.downloads_folder)
    except Exception as e:
        input('Error encountered loading ZACI ADIR dataframe.\n' + str(e))
        exit()
    try:
        dx, dme, credit_hold = ltp.merge_zaci_dataframes(zaci_adir_df, zaci_adus_df, Dir.zaci_folder)
    except Exception as e:
        input('Error encountered merging ZACI dataframes.\n' + str(e))
        exit()

    # Create dataframes for the Provisioning Reports by calling the LoadToPandas function  
    # PH Status Report
    print('Loading Provisioning Status report...')
    try:
        ph_status_returns = ltp.ph_status_dataframe(Dir.prov_excel, Dir.ph_status_file)
        ph_status_dataframes = ph_status_returns[:5]
        JOIN = ph_status_returns[5]
    except Exception as e:
        input('Error encountered loading PH Status dataframe.\n' + str(e))
        exit()

    # PH Aging Report
    try:
        print('Loading Provisioning Aging report...')
        ph_aging_dataframes = ltp.ph_aging_dataframe(Dir.prov_excel, Dir.ph_aging_file, JOIN)
    except Exception as e:
        input('Error encountered loading PH Aging dataframe.\n' + str(e))
        exit()
    
    # Create emtpty list to append dataframe metrics. This will be sent as a parameter to the send email function
    metrics_for_email = []

    # Create dataframes for the BART Reports by calling the LoadToPandas function 
    print('Loading BART Error report...')
    try:
        bart_error_df, bart_error_orders = ltp.bart_dataframe(Dir.bart_error_file)
        metrics_for_email.append(bart_error_orders)
    except Exception as e:
        input('Error encountered loading BART Error dataframe.\n' + str(e))
        exit()

    print('Loading BART Duplicate report...')
    try:
        bart_duplicate_df, bart_duplicate_orders = ltp.bart_dataframe(Dir.bart_duplicate_file)
        metrics_for_email.append(bart_duplicate_orders)
    except Exception as e:
        input('Error encountered loading BART Duplicate dataframe.\n' + str(e))
        exit()

    try:
        print('Loading BART No Provisioning report...')
        bart_no_provision_df, bart_no_prov_orders = ltp.bart_dataframe(Dir.bart_no_provision_file)
        metrics_for_email.append(bart_no_prov_orders)
    except Exception as e:
        input('Error encountered loading BART No Provisioning dataframe.\n' + str(e))
        exit()
    
    # Create dataframes for the VFX3 Reports by calling the LoadToPandas function 
    print('Loading VFX3 IHC report...')
    try:
        vfx3_ihc_df, vfx3_ihc_count = ltp.vfx3_dataframe( Dir.vfx3_ihc_file)
        metrics_for_email.append(vfx3_ihc_count)
    except Exception as e:
        input('Error encountered loading VFX3 IHC dataframe.\n' + str(e))
        exit()

    try:
        print('Loading VFX3 ADIR report...')
        vfx3_adir_df, vfx3_adir_count = ltp.vfx3_dataframe(Dir.vfx3_adir_file)
        metrics_for_email.append(vfx3_adir_count)
    except Exception as e:
        input('Error encountered loading VFX3 ADIR dataframe.\n' + str(e))
        exit()

    try:
        print('Loading VFX3 ADUS report...')
        vfx3_adus_df, vfx3_adus_count = ltp.vfx3_dataframe(Dir.vfx3_adus_file)
        metrics_for_email.append(vfx3_adus_count)
    except Exception as e:
        input('Error encountered loading VFX3 ADUS dataframe.\n' + str(e))
        exit()

    # Create dataframe for the V_UC Report by calling the LoadToPandas function 
    print('Loading V_UC report...')
    try:
        v_uc_df, v_uc_email_df = ltp.vuc_dataframe(Dir.v_uc_file) 
        metrics_for_email.append(v_uc_email_df)
    except Exception as e:
        input('Error encountered loading VUC dataframe.\n' + str(e))
        exit()

    # Create dataframe for the ZISXERROR Report by calling the LoadToPandas function 
    print('Loading ZISXERROR report...')
    try:
        zisxerror_df, zisxerror_orders = ltp.zisexerror_dataframe(Dir.zisxerror_file) 
        metrics_for_email.append(zisxerror_orders)
    except Exception as e:
        input('Error encountered loading ZISXERROR dataframe.\n' + str(e))
        exit()

     # Create dataframe for the ZISXERROR Report by calling the LoadToPandas function 
    print('Loading P_Status report...')
    try:
        p_status_df, p_status_order = ltp.p_status_dataframe(Dir.p_status_file) 
        metrics_for_email.append(p_status_order)
    except Exception as e:
        input('Error encountered loading P_Status dataframe.\n' + str(e))
        exit()

    # archive previously ran files before wrting out new files
    archive_reports(Dir.zaci_excel, Dir.zaci_archive)
    archive_reports(Dir.prov_excel, Dir.prov_archive)
    archive_reports(Dir.credit_hold_file, Dir.zaci_archive)

    # Write the various dataframes to excel by calling the WriteToFile function
    print('\nWriting reports out to Excel...')
    wtf.zaci_to_excel(dx, dme, credit_hold, Dir.zaci_folder)
    wtf.ph_to_excel(ph_status_dataframes, ph_aging_dataframes, Dir.provisioning_folder)
    wtf.excel(bart_error_df, Dir.output_folder, Dir.bart_error_file[72:-4] + '.xlsx') 
    wtf.excel(bart_duplicate_df, Dir.output_folder, Dir.bart_duplicate_file[72:-4] + '.xlsx')
    wtf.excel(bart_no_provision_df, Dir.output_folder, Dir.bart_no_provision_file[72:-4] + '.xlsx')
    wtf.excel(vfx3_ihc_df, Dir.output_folder, Dir.vfx3_ihc_file[72:-4] + '.xlsx')
    wtf.excel(vfx3_adir_df, Dir.output_folder, Dir.vfx3_adir_file[72:-4] + '.xlsx')
    wtf.excel(vfx3_adus_df, Dir.output_folder, Dir.vfx3_adus_file[72:-4] + '.xlsx')
    wtf.excel(v_uc_df, Dir.output_folder, Dir.v_uc_file[72:-4] + '.xlsx') 
    wtf.excel(zisxerror_df, Dir.output_folder, Dir.zisxerror_file[72:-4] + '.xlsx')
    wtf.excel(p_status_df, Dir.output_folder, Dir.p_status_file[72:-4] + '.xlsx')

    # Pass the provisioning dataframes into DataframeLength function to count the number of entries
    ph_status_counts = Dataframe_Length.count(ph_status_dataframes)
    metrics_for_email.append(ph_status_counts)
    ph_aging_counts = Dataframe_Length.count(ph_aging_dataframes)
    metrics_for_email.append(ph_aging_counts)

    # Call send_emails function
    send_emails(metrics_for_email)

###########################################################################
def send_emails(metrics_for_email):

    #ldap = 'grwillia'
    credit_hold_file = Dir.credit_hold_file
    provisioning_file = Dir.prov_excel
    
    # Pass in the dataframe metrics and call the email function
    html = HtmlEmail.html_for_email(metrics_for_email)
    QEndEmail.send_email(html, ldap, provisioning_file) 
    QEndEmail.send_credit_hold_email(ldap, credit_hold_file)

###########################################################################
if __name__ == "__main__":
    LogIn()
