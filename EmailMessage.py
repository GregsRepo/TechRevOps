
import win32com.client as win32
import pandas as pd

# Need varibales that count the number of entries in the various VFX reports
# Need variables that count the New, Booking Complete, Prov In Progress, and Prov Errors in the provisioning reports
# Need a varibale that holds the path to the ZACI report
# Need variable to hold dataframe from V_UC Report. This might be converted to html
# Need varibale to hold output from ZISXERROR_Report. Might be a string that contains the order and error message
# Need variables for BART Error reports. Not sure if this will be seperate for each report or just a count of errors across the 3 reports
message = """




V_UC Report 
3 Entries & 1 ZAV
40159460   Contract
    50 Configuration
    60 Configuration
    70 Configuration

60890830   Credit
       Order reason (reason for the business transaction)
       Sales Office
       Sales Group

145280620  Order
    10 Subtotal 1 from pricing procedure for condition
    10 Net value of the order item in document currency

VFX3 Report 
D001 - 1 Entries
0001 - 0 Entries 
I001 - 0 Entries

ZISXERROR Report 
1 for ADUS DX 40158623 ADCL SEARCH:OD TECH FEE Overlap

Provisioning Status Report  (see attached PH_Status_Report)
The report was run for Q1.
New
5 orders in new status
Booking Complete
16 orders in BC status
Provisioning In Progress
78 orders in PP status
Provisioning Error
0 errors

Provisioning Aging Report (see attached PH_Status_Report)
The report was run for the fiscal year.
New
7 deals in new status
Booking Complete
20 orders in BC status
Provisioning In Progress
104 deals in PP status
Provisioning Error
0 errors

BART Error Report 
The report was run for the Q1.
ADUS - 1
ADIR - 0

P Report
No Entries for Q1
"""

html= """\
<html>
    <head>
        <style>
            tr:hover {background-color:yellow;}
        </style>
    </head>
      <body>
           <p>Hi all,
           <br>
           <br>
            Please see below findings from QE Reports as of 02/27/2020
           <br>  
           <br> 
            Correspondence Monitoring for CI Invoices
            <br> 
            ADUS - 0 Items
            <br>
            ADIR - 0 items 
            <br>
            <br>
            Billing document/s not Invoiced in PRD 
            <br>
            Clear
            <br>
            <br>
            Billable BITs not Billed in PRD
            <br> 
            <br>
            Clear
            <br>
            <br>
            Raw BITs created in PRD 
            <br>
            Clear
           <br>
           <br>
            ADUS  & ADIR ZACI Billing Report
            <br>
            The report was run for Q1.
            <br>
            See attached "ResultBackground" for more details.
            <br>
            DX - Orders setup to bill from a future date + orders invoiced in billing run since report ran.
            <br>
            DME - Orders setup to bill from a future date + orders invoiced in billing run since report ran.
           </p>
       </body>
</html>
"""

# html_female = female_df.to_html(index = False)
# html_male = male_df.to_html(index = False)

# html += html_female
# html += html_male

# message2 = MIMEText(html.encode('utf-8'),'html','utf-8')