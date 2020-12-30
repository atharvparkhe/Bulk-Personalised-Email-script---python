# import the necessary libraries and modules
import smtplib
from email.message import EmailMessage
import pandas as pd
from email.message import EmailMessage
import imghdr

# your email credentials
mailid="yourmailid@gmail.com"				#	<---------- Your email id
mailpw="yourpassword"						#	<----------	Your email password

# attachment section
attachments=['Sample.pdf','meme.jpg'] 		#	<---------- Attachments
df=pd.read_excel("info.xlsx")

for mailids in df['Emailid'].values:
	c=df.loc[df['Emailid']==mailids]
	for names in c['Name'].values :
		# email body
		txt="Hi "+ str(names) + " This is a script to send multiple personalised emails(bulk email marketing).\nCheck out the attached file.\nSend feedback through the whats app group."			# <-------------- Message
	msg=EmailMessage()
	# email subject
	msg['Subject']="Test for sending personalized email with attachments" 		#  <----------  Subject
	msg['From']=mailid
	msg['To']=mailids
	msg.set_content(txt)
	for attachment in attachments:
		file_type=attachment.split('.')
		file_type=file_type[1]
		if file_type == 'jpg' or file_type == 'JPG' or file_type == 'png' or file_type == 'PNG' :
			with open(attachment, 'rb') as  f:
			      file_data=f.read()
			      img_type=imghdr.what(attachment)
			msg.add_attachment(file_data, maintype='image', subtype=img_type, filename=f.name)
		else:
			with open(attachment, 'rb') as  f:
			      file_data=f.read()
			msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=f.name)
			
	#Sending...
	try:
	    server=smtplib.SMTP_SSL("smtp.gmail.com", 465)
	    server.login(mailid, mailpw)
	    server.send_message(msg)
	    server.quit()
	    print("Successfull")
	except:
	    print("Failed")