import pandas as pd
import smtplib
import email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

book = pd.read_excel("Book1.xlsx")
email_reciever = book["Email"].values
message = MIMEMultipart()
message["From"] = "placements@iic.ac.in"
message["subject"] = "Jr. Placement Interviews"
file = open("message1.txt","r")
body = file.read()
message.attach(MIMEText(body,"plain"))
with smtplib.SMTP("smtp.gmail.com",587) as server:
	server.starttls()
	server.login("placements@iic.ac.in","Asdfzxcv1")
	for id in email_reciever : 
		del message['To']
		message["To"] = id
		text = message.as_string()
		server.sendmail("placements@iic.ac.in",id,text)