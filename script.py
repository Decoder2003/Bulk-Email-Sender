import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.message import EmailMessage

SenderAddress = "Your Email"
password = "Your password"

e = pd.read_excel("contacts.xlsx")
recipients = e['Emails'].values
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(SenderAddress, password)
html = open('template.html', encoding='utf8').read()
var = EmailMessage()

var["subject"] = "INVITATION for imAgine.py v2.1"
var["from"] = SenderAddress
var.set_content('Invitation')
var.add_alternative(html, subtype='html')

for i in range(len(recipients)):
    var['to'] = recipients[i]
    print('Sno -->',i+1,' ',var['to'])
    try:
        server.sendmail(SenderAddress, recipients[i], var.as_string())
    except:
        print('Sno -->',i+1,' ',var['to'],'::NOT SENT::Email Wrong::')
    var = MIMEText(html,'html','utf-8')
    var["subject"] = "INVITATION for imAgine.py v2.1"
    var["from"] = SenderAddress
server.quit()

# Made with love by Krishna Aggarwal & Dhruv Kalra