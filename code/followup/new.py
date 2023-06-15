import gspread
from oauth2client.service_account import ServiceAccountCredentials

#gaining access and credentials Google drive and Google sheets
from auth import *

import numpy as np
import pandas as pd
from gspread_dataframe import get_as_dataframe, set_with_dataframe

#Naming and gaining access to a specific workbook and specific sheets within that workbook.
ws1 = 'Form Responses (Do not Edit)'
ws2 = "Porsha's Leads"
ws3 = "Amy's Leads"

ws_live = client.open('Performance and Travel Form (Responses)').worksheet(ws1)
wsp = client.open('Performance and Travel Form (Responses)').worksheet(ws2)
wsa = client.open('Performance and Travel Form (Responses)').worksheet(ws3)

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

values = pd.DataFrame(ws_live.row_values(2))
values[0].replace('', np.nan, inplace=True)
values.dropna(subset=[0], inplace=True)

i = 0 
index = np.array([0])
while i < len(values.index)-1:
    i += 1
    index = np.append(index, i)
values = values.set_index(index)

print(values)

school_name_row = input("What line is the school name on? (default is line 11) \n")

if school_name_row == "":
    school = values.iat[11,0]
else:
    while int(school_name_row) > int(len(values.index)) or int(school_name_row) < 1: 
        print("Input is out of range\n")
        school_name_row = input("What line is the school name on? (default is line 12) \n")

    school = str(values.iat[int(school_name_row),0])

print(values)
