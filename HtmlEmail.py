
import datetime as dt

# Get todays date as a datetime object 
today = dt.date.today()
date = today.strftime('%m/%d/%Y') 

##########################################################################################
def html_for_email(metrics_for_email):

    # Reassign the email metrics back to variables for readability
    bart_error_orders = metrics_for_email[0]
    bart_duplicate_orders = metrics_for_email[1]
    bart_no_prov_orders = metrics_for_email[2]
    vfx3_ihc_count = metrics_for_email[3]
    vfx3_adir_count = metrics_for_email[4]
    vfx3_adus_count = metrics_for_email[5]
    v_uc_email_df = metrics_for_email[6]
    zisxerror_orders = metrics_for_email[7]
    ph_status_counts = metrics_for_email[8]
    ph_aging_counts = metrics_for_email[9]

    ph_aging_bc = ph_aging_counts[0]
    ph_aging_pip = ph_aging_counts[1]
    ph_aging_new = ph_aging_counts[2]
    ph_aging_pe = ph_aging_counts[3]

    ph_status_bc = ph_status_counts[0]
    ph_status_pip = ph_status_counts[1]
    ph_status_new = ph_status_counts[2]
    ph_status_pe = ph_status_counts[3]

    # vfx3_ihc_count, vfx3_adir_count, vfx3_adus_count

    html_table = v_uc_email_df.to_html(index = False) 
    html_table1 = zisxerror_orders.to_html(index = False)
    html_table2 = bart_error_orders.to_html(index = False) 
    html_table3 =  bart_no_prov_orders.to_html(index = False) 
    html_table4 = bart_duplicate_orders.to_html(index = False) 
    

    html = """\
    <html>
    <head></head>
    <body>
        Hi all,
        <br>
        <br>
            Please see below findings from QE Reports as of """+date+"""
        <br>  
        <br> 
            Correspondence Monitoring for CI Invoices
        <br> 
            ADUS - ?? Items
        <br>
            ADIR - ?? items 
        <br>
        <br>
            Billing document/s not Invoiced in PRD 
        <br>
            ??????
        <br>
        <br>
            Billable BITs not Billed in PRD
        <br>
            ??????
        <br>
        <br>
            Raw BITs created in PRD 
        <br>
            ??????
        <br>
        <br>
            ADUS  & ADIR ZACI Billing Report
        <br>
            The report was run for Q?.
        <br>
            Click this link to access the file - 
        <a href="\\\\du1isi0\\order_management\\TechRevOps\\Reports\\Automated\\ZACI\\ZACI_Report.xlsx">ZACI_Report.xlsx</a>
        <br>
            DX - Orders setup to bill from a future date + orders invoiced in billing run since report ran.
        <br>
            DME - Orders setup to bill from a future date + orders invoiced in billing run since report ran.
        <br>
        <br>
            V_UC Report 
    </body>
    </html>
    """
    html2 = """\
    <html>
    <head></head>
    <body>
        <br>
            VFX3 Report 
        <br>
            D001 - """ +vfx3_adir_count+ """ Entries
        <br>
            0001 - """ +vfx3_adus_count+ """ Entries 
        <br>
            I001 - """ +vfx3_ihc_count+  """ Entries
        <br>
        <br>
            ZISXERROR Report 
    </body>
    </html>
    """ 
    html3 = """\
    <html>
    <head></head>
    <body>
        <br>
            Provisioning Status Report  (see attached PH_Status_Report)
        <br>
            The report was run for ????????.
        <br>
            New - """ +ph_status_new+ """ orders in new status
        <br>
            Booking Complete - """ +ph_status_bc+ """ orders in BC status
        <br>
            Provisioning In Progress - """ +ph_status_pip+ """ orders in PP status
        <br>
            Provisioning Error - """ +ph_status_pe+""" errors
        <br>
        <br>
            Provisioning Aging Report (see attached PH_Aging_Report)
        <br>
            The report was run for ?????????.
        <br>
            New - """ +ph_aging_new+ """ deals in new status
        <br>
            Booking Complete - """ +ph_aging_bc+ """ orders in BC status
        <br>
            Provisioning In Progress -""" +ph_aging_pip+ """ deals in PP status
        <br>
            Provisioning Error - """ +ph_aging_pe+ """ errors
    </body>
    </html>
    """
    html4 = """\
    <html>
    <head></head>
    <body>
        <br>
            BART Error Report 
        <br>
            The report was run for the Q1.
    </body>
    </html>
    """

    html5 = """\
    <html>
    <head></head>
    <body>
        <br>
            BART No Provisioning Report 
        <br>
            The report was run for the Q1.
    </body>
    </html>
    """

    html6 = """\
    <html>
    <head></head>
    <body>
        <br>
            BART Duplicate Report 
        <br>
            The report was run for the Q1.
    </body>
    </html>
    """

    html7 = """\
    <html>
    <head></head>
    <body>
        <br>
            P Report
        <br>
            ?? Entries for Q1
    </body>
    </html>
    """
    

    
    html += html_table 
    html += html2
    html += html_table1
    html += html3
    html += html4
    html += html_table2
    html += html5
    html += html_table3
    html += html6
    html += html_table4
    html += html7

    return html

##########################################################################################