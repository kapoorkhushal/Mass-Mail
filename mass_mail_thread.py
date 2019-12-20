import email
import smtplib
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
import email.mime.application
import threading

sender_email = "<sender-id>"
subject = "<subject for mail>"
message = MIMEMultipart()
message["From"] = sender_email
message["subject"] = subject
file = open("message.txt","r")									#file with text to be sent
body = file.read()
message.attach(MIMEText(body,"plain"))

attachment1 = "<attachment>"
attachment2 = "<attachment>"
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


def function(file_name):
	book = pd.read_excel(file_name)
	reciever_email = book['<column name for Email-Id>'].values
	with smtplib.SMTP("smtp.gmail.com",587) as server:
		server.starttls()
		server.login("<login - id>","<password>")
		for id in reciever_email:
			del message['To']
			message["To"] = id
			text = message.as_string()
			try:
				server.sendmail(sender_email,id,text)
			except:
				print(id+'\n')
	print("Success - "+file_name)

if __name__ == '__main__':
	t1 = threading.Thread(target = function,args=('<filename.xlsx>',))
	t2 = threading.Thread(target = function,args=('<filename.xlsx>',))
	t1.start()
	t2.start()
	t1.join()
	t2.join()
	print("Done")
