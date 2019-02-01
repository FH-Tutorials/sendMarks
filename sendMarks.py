import pandas as pd
import configparser as cf
import getpass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import argparse
import sys
import os.path

default_config='setup.ini'

# retrieve excel handle if possible otherwise
# provide parser error
def get_excel_handle(parser, arg):
    if not os.path.exists(arg):
        parser.error("The file %s does not exist!" % arg)
    else:
        try:
            return pd.ExcelFile(arg)
        except Exception as e:
            parser.error("Error processing file %s: %s" % (arg, str(e)))

# setup argument parser
parser = argparse.ArgumentParser(description='Process Excel File')
parser.add_argument('xlsfile', help='path to excel file', type=lambda x:get_excel_handle(parser, x))
parser.add_argument('--config', help='path to config file (default: %s)' % default_config)
parser.add_argument('--login', help='enable SMTP login', action='store_true')
parser.add_argument('--starttls', help='enable starttls', action='store_true')
parser.set_defaults(config=default_config,login=False,starttls=False)
args = parser.parse_args()

# read config file
config = cf.ConfigParser()
config.read(args.config)

# pick up excel handle
xl = args.xlsfile

# connect to STMP server
connectString = config['MAIL']['smtp'] + ":" + config['MAIL']['port']
print("Connecting to: %s ..." % connectString)
server = smtplib.SMTP(connectString)

# enable starttls
if args.starttls: server.starttls()

# login using credentials 
if args.login:
    print("Logging in with User %s: " % config['MAIL']['user'])
    password = getpass.getpass()
    server.login(config['MAIL']['user'], password)

# Load first sheet into a pandas-DataFrame
df = xl.parse(xl.sheet_names[0])

print("sending messages...")
# iterate excel rows
for index, row in df.iterrows():
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
