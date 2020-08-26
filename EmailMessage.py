
import win32com.client as win32
#import LoadToPandas as lpt
import pandas as pd


# Need varibales that count the number of entries in the various VFX reports
# Need variables that count the New, Booking Complete, Prov In Progress, and Prov Errors in the provisioning reports
# Need a varibale that holds the path to the ZACI report
# Need variable to hold dataframe from V_UC Report. This might be converted to html
# Need varibale to hold output from ZISXERROR_Report. Might be a string that contains the order and error message
# Need variables for BART Error reports. Not sure if this will be seperate for each report or just a count of errors across the 3 reports
message = """







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
            <br>
            <br>
            V_UC Report 
            "+html_table+" 
           </p>
       </body>
</html>
"""

# html_female = female_df.to_html(index = False)
# html_male = male_df.to_html(index = False)

# html += html_female
# html += html_male

# message2 = MIMEText(html.encode('utf-8'),'html','utf-8')