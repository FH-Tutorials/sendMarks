# Import pandas
import pandas as pd
import configparser as cf

# read config file
config = cf.ConfigParser()
config.read("setup.ini")

# output some infos
print("Using:")
print("SMTP " + config['MAIL']['smtp'])
print("User " + config['MAIL']['user'])

# Assign spreadsheet filename to `file`
file = "./Results.xlsx"

# Load spreadsheet
xl = pd.ExcelFile(file)

# Load first sheet into a pandas-DataFrame
df1 = xl.parse(xl.sheet_names[0])

# iterate excel rows
for index, row in df1.iterrows():
    print('---')
    print(config.get('TEXT_DE', 'text', vars=row.to_dict()))
