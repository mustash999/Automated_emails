'''
 # @ ------------------------------------------------------------------------------: @ #

 # @ Author: 							Mustash										 @ #
 # @   : 																			:@ #
 # @ Create Time: 						23-09-20 21:50								 @ #
 # @ Modified time: 					24-10-31 02:33								 @ #
 # @ Modified by: 						Mustash										 @ #
 # @    : 																			:@ #
 # @ web: 								www.geniuspandatech.com						 @ #
 # @ Description: 						Put the description here
 # @ - ----------------------------------------------------------------------------: @ #

 '''

import									smtplib
from email.mime.multipart	import		MIMEMultipart
from email.mime.text		import		MIMEText
from email.mime.image		import		MIMEImage
from getpass				import 		getpass
import pandas				as			pd
import									os
from dotenv					import		load_dotenv
import									ms_compose_email

# Email configuration
load_dotenv() ########  Attention --- >  see Genius Panda panda video on this  

sender_name		= os.getenv('NAME')
sender_email	= os.getenv('EMAIL')
company			= os.getenv('COMPANY')
filename		= os.path.join( os.getcwd(), 'res',  os.getenv('FILENAME'))
sernder_server	= os.getenv('SERVER')
send_port		= os.getenv('PORT')


def get_group(filename):
	df = pd.read_excel(filename)
	recievers= df['Email'].tolist()
	return recievers

receivers_emails = get_group(filename)
subject = "Merry Christmas and a Happy New Year"

# Get the sender's password securely (you can also hardcode it)
password = os.getenv('PASSWORD')

# Establish a secure SMTP connection with the server
server_connection = None
try: # this is for ttls connection (for ssl  there is another way)
	server_connection = smtplib.SMTP(sernder_server, send_port)
	server_connection.starttls()  
	server_connection.login(sender_email, password)

	# Loop through recipients and send individual emails
	for email in receivers_emails:
		msg = MIMEMultipart()
		msg["From"] = f"{sender_name} <{sender_email}>"
		msg["To"] = email
		msg["Subject"] = subject
		with open('msgs_to_send/' + email + '.txt', 'r', encoding='utf-8' ) as file:
			message = file.read()
		msg.attach(MIMEText(message, "plain"))

		image_path = os.path.join(os.getcwd(), 'res', 'card.jpg')
		with open(image_path, 'rb') as image_file:
			image = MIMEImage(image_file.read())
			image.add_header('Content-Disposition', 'attachment', filename= image_path)
		msg.attach(image)


		server_connection.sendmail(sender_email, email, msg.as_string())
		print(f"Email sent successfully to {email}")

except Exception as e:
	print(f"An error occurred: {str(e)}")

finally:
	if server_connection:
		server_connection.quit()

