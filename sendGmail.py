""" 
	A Python 3.x script to send emails through Gmail based on user input used to fill in a template email body.
	This example also adds one .pdf attachment file to the email being sent.
	Karen Li @ karenli.co 
"""

### IMPORT REQUIRED LIBRARIES ###

import smtplib
import getpass 
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import numpy as np

### DEFINE PARAMETERS ###

# TODO: modify these
yourName = "Karen Li"
yourMajor = "Computer Science"
yourCCAddrs = ["karenli.614@gmail.com", "karen.li@ucla.edu"]

COMMASPACE = ', '
# email subject and source path for the files
subject = "UCLA Startup Fair 2017 Invitation"
# template file and attachment file are inside the same folder as this script
# the file path and file name are the same
newTemplateFilePath = "new-template.txt";
returningTemplateFilePath = "returning-template.txt";
attachmentFilePath = "Sponsorship-Brochure.pdf"

### CREATE STRINGS FOR TEMPLATES ###

newTemplateFile = open(newTemplateFilePath, 'r')
# read the contents of the file into a string
newTemplateString = newTemplateFile.read()
newTemplateFile.close()

returningTemplateFile = open(returningTemplateFilePath, 'r')
# read the contents of the file into a string
returningTemplateString = returningTemplateFile.read()
returningTemplateFile.close()

### GMAIL ACCOUNT LOG-IN ###

print("=== Gmail Account Log-in ===")
username = input("Please enter your email address: ")
password = getpass.getpass("Please enter your password: ")

### LOAD EMAIL DATA FROM CSV ###

csvFilePath = "tester.csv"
# ignore the first 10 rows of spreadsheet since they aren't company email info; toggle this value if needed
numRowsToSkip = 10
# get indices for [Company Name, Contact Email, Status] columns
columnIndices = (0, 2, 7)

# get 2D array with a row for each company
contactInfo = np.genfromtxt(csvFilePath, delimiter = ',', skip_header = numRowsToSkip, usecols = columnIndices, dtype='str')

### SEND EMAILS ###

for contact in contactInfo:
	# TODO: remove after testing
	print(contact)

	companyName = contact[0].strip()
	companyEmail = contact[1].strip()
	companyStatus = contact[2].strip()

	# skip if email address is not valid (doesn't contain '@' character)
	if not re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", companyEmail):
		continue

	# fill in correct template based on status of company
	if companyStatus == "Committed":
		body = returningTemplateString.format(yourName, yourMajor)
	else:
		body = newTemplateString.format(yourName, yourMajor, companyName)

	# create the container (outer) email message and add the body of the email to the email container
	msg = MIMEMultipart()
	msg['From'] = username
	# format the list of 'To:' and 'CC:' addresses into strings where each address is separated by a comma
	msg['To'] = companyEmail
	msg['Cc'] = COMMASPACE.join(yourCCAddrs)
	msg['Subject'] = subject

	# add the body of the email to the email container
	msg.attach(MIMEText(body, 'plain'))

	# attach the sponsorship brochure to the email container
	# open attachment file in binary mode since it isn't a text file
	attachmentFile = open(attachmentFilePath, 'rb')
	attachment = MIMEApplication(attachmentFile.read(), _subtype="pdf")
	attachmentFile.close()
	# need to add header  with info to specify that attachment should not be displayed right away
	attachment.add_header('Content-Disposition','attachment',filename=attachmentFilePath)
	# attach the file to the email container
	msg.attach(attachment)

	# use a secure (SSL) connection to connect with the SMTP server to send the email
	# if there's an error, print out the exception error message
	try:
		server_ssl = smtplib.SMTP_SSL('smtp.gmail.com', 465)
		# initiate SMTP conversation with SMTP server
		server_ssl.ehlo()
		# provide log-in info
		server_ssl.login(username, password)
		# flatten out the email object into actual text
		text = msg.as_string()
		# send the email (need to specify the sender, a list of recipients, and the email text)
		server_ssl.sendmail(username, [companyEmail]+yourCCAddrs, text)
		server_ssl.close()
		print("Email sent successfully!")
	except Exception as e:
		print(e)
		print("Could not send email!")