import smtplib
import email
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
import email.mime.application

book = pd.read_excel("Thread_Revive.xlsx",'2016')
reciever_email = book['E-mail ID'].values
sender_email = "placements@iic.ac.in"
subject = "Invitation for Placement Drive - 2020"
message = MIMEMultipart()
message["From"] = sender_email
message["subject"] = subject
file = open("message.txt","r")
body = file.read()
message.attach(MIMEText(body,"plain"))

attachment1 = "syllabus.pdf"
attachment2 = "IIC PLACEMENT BROCHURE-compressed.pdf"
with open(attachment1,"rb") as attachment:
	part = MIMEBase("application", "octet-stream")
	part.set_payload(attachment.read())

encoders.encode_base64(part)
part.add_header("Content-Disposition","attachment", filename= attachment1)
message.attach(part)

with open(attachment2,"rb") as attachment:
	part = MIMEBase("application", "octet-stream")
	part.set_payload(attachment.read())

encoders.encode_base64(part)
part.add_header("Content-Disposition","attachment", filename= attachment2)
message.attach(part)

#server = smtplib.SMTP("smtp.gmail.com",587)
#server.starttls()
#server.login("khushal.kapoor@iic.ac.in","")
#body = "subject : "+subject+"\n\n"+message.format(subject,message)

with smtplib.SMTP("smtp.gmail.com",587) as server:
	server.starttls()
	server.login("placements@iic.ac.in","Asdfzxcv1")
	for id in reciever_email:
		del message['To']
		message["To"] = id
		text = message.as_string()
		try:
			server.sendmail(sender_email,id,text)
		except:
			print(id+'\n')
print("Success")