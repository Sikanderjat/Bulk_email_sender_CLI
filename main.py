import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


data =pd.read_excel("data.xlsx") #getting data from bool1.xlsx file
print(data)
email_address=data.get("email") # getting all emails 
# print(email_address)
data_list=list(email_address) # store emails into a list

print(data_list)

try:
    server =smtplib.SMTP("smtp.gmail.com",587)
    server.starttls()
    server.login("your email address","app password") # ex: "yiys jrlh lauw hfmd"
    from_ ="your email address"
    To_=data_list
    msg = MIMEMultipart()
    msg['From'] = from_
    msg['subject']="this is testing email"
    
    html='''
    <html>
    <head></head>

    <body>
        <h1>This email is sending by sikander</h1>
    </body>
    </html>    
    '''
    text=MIMEText(html,"html")
    msg.attach(text)

    server.sendmail(from_,To_,msg.as_string())
    print("email send successfully")
except Exception as e:
    print(e)
finally:
    server.quit()
