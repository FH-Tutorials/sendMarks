# Import pandas
import pandas as pd
import configparser as cf
#import getpass #https://pymotw.com/2/getpass/

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import argparse
import sys
import os.path

default_config='setup.ini'
def get_excel_handle(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
    else:
        try:
            return pd.ExcelFile(arg)
        except Exception as e:
            parser.error("Error processing file %s: %s" % (arg, str(e)))

parser = argparse.ArgumentParser(description='Process Excel File')
parser.add_argument('xlsfile', help='path to excel file', type=lambda x:get_excel_handle(parser, x))
parser.add_argument('--config', help='path to config file (default: %s)' % default_config)
parser.set_defaults(config=default_config)
args = parser.parse_args()

# read config file
config = cf.ConfigParser()
config.read(args.config)


#password = getpass.getpass()

# Assign spreadsheet filename to `file`
file = args.xlsfile

# pick up handle
xl = args.xlsfile

# Load spreadsheet

# output some infos
print("Using:")
print("SMTP " + config['MAIL']['smtp'])
print("User " + config['MAIL']['user'])

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
