# Import pandas
import pandas as pd
import configparser as cf
import getpass #https://pymotw.com/2/getpass/

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


#todo: mail
#  https://code.tutsplus.com/tutorials/sending-emails-in-python-with-smtp--cms-29975

# read config file
config = cf.ConfigParser()
config.read("setup.ini")

# output some infos
print("Using:")
print("SMTP " + config['MAIL']['smtp'])
print("User " + config['MAIL']['user'])

#password = getpass.getpass()

# Assign spreadsheet filename to `file`
file = "./Results.xlsx"

# Load spreadsheet
xl = pd.ExcelFile(file)

# Load first sheet into a pandas-DataFrame
df1 = xl.parse(xl.sheet_names[0])

connectString = config['MAIL']['smtp'] + ":" + config['MAIL']['port']
print(connectString)

# create server
server = smtplib.SMTP(connectString)

# server.starttls()

# Login Credentials for sending the mail
# server.login(config['MAIL']['user'], password)

print("sending messages...")
# iterate excel rows
for index, row in df1.iterrows():
    # create message object instance
    msg = MIMEMultipart()
    msg['From'] = config['MAIL']['From']
    msg['Reply-To'] = config['MAIL']['ReplyTo']
    msg['CC'] = config['MAIL']['CC']
    msg['To'] = row[config['MAIL']['addresscolumn']]
    msg['Subject'] = config['MAIL']['Subject']

    # add in the message body
    msg.attach(MIMEText(config.get('TEXT_DE', 'text', vars=row.to_dict()), 'plain'))

    # send the message via the server.
    server.sendmail(msg['From'], msg['To'], msg.as_string())

# done
server.quit()