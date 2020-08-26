import smtplib, sys
import EmailMessage 
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime as dt

# Get todays date as a datetime object 
today = dt.date.today()
date = today.strftime('%m/%d/%Y') 

def send_email(html, ldap): 

    '''For trying to attach files later, try:
    https://stackoverflow.com/questions/37204979/python-how-do-i-send-multiple-files-in-the-email-i-can-send-1-file-but-how-to-s
    https://realpython.com/python-send-email/#adding-attachments-using-the-email-package
    https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
    '''
    
    print('\nSending email to Stakeholders...')
    try:
        server = smtplib.SMTP('namail.corp.adobe.com', 25)
        msg = MIMEMultipart()
        msg['Subject'] = 'Quarter End Reports - ' + date
        msg['From'] = 'DoNotReply@adobe.com'
        msg['To'] = ldap + '@adobe.com'
        # body = MIMEText(message) # convert the body to a MIME compatible string
        part1 = MIMEText(html, 'html')
        #part2 = MIMEText(html2, 'html')
        msg.attach(part1)
        #msg.attach(part2)
        server.send_message(msg)
        server.sendmail
        server.quit()
    except:
        sys.exc_info()
        print('Error sending mail')
        exit()

##########################################################################################
