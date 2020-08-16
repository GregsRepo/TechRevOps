import smtplib, ssl
import EmailMessage as EM

port = 465  # For SSL
smtp_server = "smtp.gmail.com"
sender_email = "gw48957@gmail.com"  # Enter your address
receiver_email = "grwillia@adobe.com"  # Enter receiver address
#password = input("Type your password and press enter: ")
password = 'Kilclare55#'
message = "\
Subject: Hi there " + EM.message

context = ssl.create_default_context()
with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message)


