# Import pandas
import pandas as pd
import configparser as cf

config = cf.ConfigParser()
config.read("setup.ini")
print(config.sections())
print(config['MAIL']['smtp'])

print(config.get('TEXT_DE', 'text', raw=True))

print(config.get('TEXT_DE', 'text', vars={
    'Vorname': 'hugo',
    'Nachname': 'meier',
    'Gesamt': '39',
    'Note': '5',
    'test': 'sss'
}))

#print(config['TEXT_DE']['text'])

# Assign spreadsheet filename to `file`
file = "set file path"

# Load spreadsheet
xl = pd.ExcelFile(file)

# Print the sheet names
print(xl.sheet_names)

# Load a sheet into a DataFrame by name: df1
df1 = xl.parse('Worksheet')

# iterate excel rows
for index, row in df1.iterrows():
    print('---')
    print(config.get('TEXT_DE', 'text', vars=row.to_dict()))
