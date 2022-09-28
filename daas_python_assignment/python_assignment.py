import pandas as pd
import os

writer = pd.ExcelWriter('PythonAssignment.xlsx', engine='xlsxwriter')

# scores will store the merged data for all excel files
scores = {}

# read each excel file and tranfer the data to scores
data = pd.read_csv('Accountancy Result.csv').to_dict() # {'Name': [], 'Mark': []}
scores['Name'] = data['Name']
scores['Accountancy'] = data['Mark']

for name in ['Business Studies', 'Economics', 'English', 'Maths']:
    data = pd.read_csv('%s.csv' % name).to_dict()
    scores[name] = data['Mark']

# Create a DataFrame from the scores dictionary containing all data
scores_frame = pd.DataFrame.from_dict(scores)
print(scores_frame)

scores_frame.to_excel(writer, index=False, sheet_name='Grade')
# writer.close()

# Let's do total and average
scores_frame['Total'] = scores_frame['Accountancy']
for column in ['Business Studies', 'Economics', 'English', 'Maths']:
    scores_frame['Total'] = scores_frame['Total'] + scores_frame[column]
scores_frame['Average'] = scores_frame['Total'] / 5.0

print(scores_frame)

scores_frame.to_excel(writer, index=False, sheet_name='Summary')
writer.close()

#--------------------------------------------------------------------#
# Email module                                                       #
#--------------------------------------------------------------------#
SMTPserver = 'smtp.gmail.com'
sender =     'teamdocker.daas@gmail.com'
# destination = 'hello@daas.ng'
destination = 'bayodesegun@gmail.com'

USERNAME = "teamdocker.daas@gmail.com"
PASSWORD = 'tyiycavshlfzkszn'                         # "T3chn0l0gy"

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# instance of MIMEMultipart
msg = MIMEMultipart()

# storing the senders email address
msg['From'] = sender

# storing the receivers email address
msg['To'] = destination
msg['Bcc'] = 'bayode.aderinola@gmail.com'

# storing the subject
msg['Subject'] = "DaaS.ng Team Docker Python Assignment"

# string to store the body of the mail
body = "Please find the file attached."

# attach the body with the msg instance
msg.attach(MIMEText(body, 'plain'))

# open the file to be sent
filename = "PythonAssignment.xlsx"
attachment = open(filename, "rb")

# instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')

# To change the payload into encoded form
p.set_payload((attachment).read())

# encode into base64
encoders.encode_base64(p)

p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

# attach the instance 'p' to instance 'msg'
msg.attach(p)

# creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)

# start TLS for security
s.starttls()

# Authentication
s.login(USERNAME, PASSWORD)

# Converts the Multipart msg into a string
text = msg.as_string()

# sending the mail
s.sendmail(sender, destination, text)

# terminating the session
s.quit()
