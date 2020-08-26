import functools, random, smtplib, string, sys
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

For improvements, you want to look at sending an e-amil seperately for the credit file and attaching files to 
both the quarter end report and credit e-mail

Might also look at general excpetion handling as bugs are discovered'''

###############################################################################################################################
def LogIn():
    # Global value to take the users ldap at login and pass it to the email fucntion at end of this script
    global ldap
    ldap = input('Please enter your LDAP or enter 0 to exit: ').lower()
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
            print('Incorrect Token. Please try again or press 0 to exit:')
            LogIn()
    else:
        print('Access Denied. Please contact Grp-TechRevOps@adobe.com for privilages.')
        LogIn()

###############################################################################################################################
def get_menu():
    
    option = input('''\nChoose an option(Pick a number. Or '0' to exit):
    0: Exit
    1: Full Quarter End Report (includes download of new SAP Reports)
    2: Reload output files (just reload files from previous SAP downloads)
    Option: ''')

    if option == "0":
        sys.exit()
    elif option == "1":
        run_sap_reports()
    elif option == "2":
        call_load_to_pandas()

    menu = ["1", "2"]
    while option not in menu:
        option = input('\nPick a number between 1 & 2: ')

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

    call_load_to_pandas()

###########################################################################
def call_load_to_pandas():

    '''Might not need the variables above. Some of the below functions are reading directly from the Directories folder
    Need to decide what is best'''
    print('\nWriting Reports to memory...:')
    # Create dataframes for the ZACI Reports by calling the LoadToPandas function 
    print('Loading ZACI report...')
    zaci_adus_df = ltp.zaci_dataframe('ADUS', Dir.downloads_folder)
    zaci_adir_df = ltp.zaci_dataframe('ADIR', Dir.downloads_folder)
    dx, dme, credit_hold = ltp.merge_zaci_dataframes(zaci_adir_df, zaci_adus_df, Dir.zaci_folder)

    # Create dataframes for the Provisioning Reports by calling the LoadToPandas function  
    # PH Status Report
    print('Loading Provisioning Status report...')
    ph_status_returns = ltp.ph_status_dataframe(Dir.downloads_folder, Dir.prov_excel, Dir.ph_status_file)
    ph_status_dataframes = ph_status_returns[:5]
    JOIN = ph_status_returns[5]
    # PH Aging Report
    print('Loading Provisioning Aging report...')
    ph_aging_dataframes = ltp.ph_aging_dataframe(Dir.downloads_folder, Dir.ph_aging_file, JOIN)
    #ph_aging_dataframes = ph_aging_returns[:4]
    
    # Create emtpty list to append dataframe metrics. This will be sent as a parameter to the send email function
    metrics_for_email = []

    # Create dataframes for the BART Reports by calling the LoadToPandas function 
    print('Loading BART Error report...')
    bart_error_df, bart_error_orders = ltp.bart_dataframe(Dir.bart_error_file)
    metrics_for_email.append(bart_error_orders)
    print('Loading BART Duplicate report...')
    bart_duplicate_df, bart_duplicate_orders = ltp.bart_dataframe(Dir.bart_duplicate_file)
    metrics_for_email.append(bart_duplicate_orders)
    print('Loading BART No Provisioning report...')
    bart_no_provision_df, bart_no_prov_orders = ltp.bart_dataframe(Dir.bart_no_provision_file)
    metrics_for_email.append(bart_no_prov_orders)
    
    # Create dataframes for the VFX3 Reports by calling the LoadToPandas function 
    print('Loading VFX3 IHC report...')
    vfx3_ihc_df, vfx3_ihc_count = ltp.vfx3_dataframe( Dir.vfx3_ihc_file)
    metrics_for_email.append(vfx3_ihc_count)
    print('Loading VFX3 ADIR report...')
    vfx3_adir_df, vfx3_adir_count = ltp.vfx3_dataframe(Dir.vfx3_adir_file)
    metrics_for_email.append(vfx3_adir_count)
    print('Loading VFX3 ADUS report...')
    vfx3_adus_df, vfx3_adus_count = ltp.vfx3_dataframe(Dir.vfx3_adus_file)
    metrics_for_email.append(vfx3_adus_count)

    # Create dataframe for the V_UC Report by calling the LoadToPandas function 
    print('Loading V_UC report...')
    v_uc_df, v_uc_email_df = ltp.vuc_dataframe(Dir.v_uc_file) 
    metrics_for_email.append(v_uc_email_df)

    # Create dataframe for the ZISXERROR Report by calling the LoadToPandas fucntion 
    print('Loading ZISXERROR report...')
    zisxerror_df, zisxerror_orders = ltp.zisexerror_dataframe(Dir.zisxerror_file) 
    metrics_for_email.append(zisxerror_orders)


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

    # Pass the provisioning dataframes into DataframeLenght function to count the number of entries
    ph_status_counts = Dataframe_Length.count(ph_status_dataframes)
    metrics_for_email.append(ph_status_counts)
    ph_aging_counts = Dataframe_Length.count(ph_aging_dataframes)
    metrics_for_email.append(ph_aging_counts)

    # Pass in the dataframe metrics and call the email function
    html = HtmlEmail.html_for_email(metrics_for_email)
    QEndEmail.send_email(html, ldap) 

###########################################################################
if __name__ == "__main__":
    LogIn()
    #get_menu()

