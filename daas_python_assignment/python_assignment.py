import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#--------------------------------------------------------------------#
# Read CSV and Write Excel file                                      #
#--------------------------------------------------------------------#
filename = "PythonAssignment.xlsx"
writer = pd.ExcelWriter(filename, engine='xlsxwriter')

# scores will store the merged data for all excel files
scores = {}

# read each excel file and tranfer the data to scores
data = pd.read_csv('Accountancy Result.csv').to_dict()   # {'Name': [], 'Mark': []}
scores['Name'] = data['Name']
scores['Accountancy'] = data['Mark']
csv_file_names = ['Business Studies', 'Economics', 'English', 'Maths']

for column_name in csv_file_names:
    data = pd.read_csv("{column_name}.csv".format(column_name=column_name)).to_dict()
    scores[column_name] = data['Mark']

# Create a DataFrame from the scores dictionary containing all data
scores_frame = pd.DataFrame.from_dict(scores)
print(scores_frame)

# Write the 'Grade' Sheet
scores_frame.to_excel(writer, index=False, sheet_name='Grade')

# Let's do total and average
scores_frame['Total'] = scores_frame['Accountancy']
for column in csv_file_names:
    scores_frame['Total'] = scores_frame['Total'] + scores_frame[column]
scores_frame['Average'] = scores_frame['Total'] / 5.0

print(scores_frame)

# Write the 'Summary' sheet and close the writer
scores_frame.to_excel(writer, index=False, sheet_name='Summary')
writer.close()

#--------------------------------------------------------------------#
# Send Excel Sheet by email                                          #
#--------------------------------------------------------------------#
USERNAME = os.environ.get('USERNAME')
PASSWORD = os.environ.get('PASSWORD')

sender = USERNAME
destination = 'hello@daas.ng'
# destination = 'bayodesegun@gmail.com'

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = destination
msg['Bcc'] = 'bayode.aderinola@gmail.com'
msg['Subject'] = "DaaS.ng Team Docker Python Assignment"
body = "Please find the file attached."

# attach the body with the msg instance
msg.attach(MIMEText(body, 'plain'))

# open the file to be sent
attachment = open(filename, "rb")
p = MIMEBase('application', 'octet-stream')
p.set_payload((attachment).read())
encoders.encode_base64(p)
p.add_header('Content-Disposition', "attachment; filename={filename}".format(filename=filename))
msg.attach(p)

# create SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)
s.starttls()
s.login(USERNAME, PASSWORD)

# send the mail
message = msg.as_string()
s.sendmail(sender, destination, message)

# terminate the session
s.quit()
