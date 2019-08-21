import smtplib, ssl
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

port = 587  # For starttls
smtp_server = "smtp.gmail.com"
print ("Please make sure you have allowed less secure apps")
print ("If you haven't please go to this link https://myaccount.google.com/lesssecureapps")
sender_email = input("Type your email and press enter: ")
password = input("Type your password and press enter: ")
wb = xl.load_workbook(r'<Replace with path of your excel sheet>')
sheet1 = wb['Sheet1']
Subject = 'Test'

names = []
emails = []

for cell in sheet1['A']:
    emails.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

context = ssl.create_default_context()
with smtplib.SMTP(smtp_server, port) as server:
    server.starttls(context=context)
    server.login(sender_email, password)
    print ("Login Succesful")
    for i in range(len(emails)):
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = names[i]
        msg['Subject'] = Subject
        text = '''
Hello {},
    This is bulk mailing test!
    {}
'''     .format(names[i],emails[i])
        msg.attach(MIMEText(text, 'plain'))
        message = msg.as_string()
        server.sendmail(sender_email, emails[i], message)
        print('Mail sent to', emails[i])
    server.quit()   

print('All emails sent successfully!')
