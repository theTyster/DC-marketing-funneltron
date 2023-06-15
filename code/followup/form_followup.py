#! python3
"""
This sheet pulls responses from the google form, condenses them into a dataframe and then 
emails the details to the people inputted and changes the color on the google sheet.
It then Increments the step.
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials

#gaining access and credentials Google drive and Google sheets
from auth import *

import numpy as np
import pandas as pd
from gspread_dataframe import get_as_dataframe, set_with_dataframe

#Naming and gaining access to a specific workbook and specific sheets within that workbook.
ws1 = 'Form Responses (Do not Edit)'
ws2 = 'Form Responses (Current)'

ws_live = client.open('Performance and Travel Form (Responses)').worksheet(ws1)
ws = client.open('Performance and Travel Form (Responses)').worksheet(ws2)

# assigns a row into a Panda data frame.
values = pd.DataFrame(ws_live.row_values(2))
requester = ws_live.acell('C2').value

# remove all empty values from data frame by replacing them with NaN and then dropping all cells with Nan value. inplace=True specifies that the cells will be replaced as opposed to appended I guess.
values[0].replace('', np.nan, inplace=True)
values.dropna(subset=[0], inplace=True)

# drop indices. So that "values" is only the raw data formatted as a list.
values = values.drop(values.index[0])

#Prints the form data into std out to make it easier to determine who the data should go to.
print(values)

from step import *
# add a column header, hide the index, and export to HTML
values.rename(columns = {0:f"<p style='color: darkgreen;'>Performance and Travel Form Response #{step}</p>"}, inplace=True)
html_table = values.style.hide_index().to_html()

#------------------

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#Establish SMTP Connection
s = smtplib.SMTP('smtp-mail.outlook.com', 587)

#Start TLS based SMTP Session
s.starttls()

#Login Using Your Email ID & Password
s.login(username, password)

#To Create Email Message in Proper Format
msg = MIMEMultipart()

## Function that determines who should be coppied on the resulting Email. And who should be the contact owner in HubSpot
def censor():
    global bigtrip

    bigtrip = input("Is this a team 3 form? (y/n/test) \n")

    global recipients

    if bigtrip == "y":
    
        recipients = ["***REMOVED***", "***REMOVED***"]
        print("Sending to Ty and Amy")
    
    elif bigtrip == "n":
    
        recipients = ["***REMOVED***", "***REMOVED***"]
        print("Sending to Ty, Porsha")

    elif bigtrip == "test":
        recipients =["***REMOVED***"]
        print("Sending you a test email")

    else:
        print("input is not recognized")
        censor()

censor()

#Get the HubSpot Id and assign the contact to the right consultant.
#Also updates a few other fields.
import hubspot
from hubspot.crm.contacts import SimplePublicObjectInput, ApiException, PublicObjectSearchRequest
import re

Amy_id = "48087941"
Porsha_id = "25967362"

try:
    if bigtrip == "n": 
        properties = {
            "lifecyclestage": "marketingqualifiedlead",
            "n2023_account_status": "Customer",
            "hubspot_owner_id": Porsha_id
        }
        regex = r"('id': )'([\d]*)"
        public_object_search_request = PublicObjectSearchRequest(filter_groups=[{"filters":[{"value":requester,"propertyName":"email","operator":"EQ"}]}])
        hs_search = str(hs_client.crm.contacts.search_api.do_search(public_object_search_request=public_object_search_request))
        result = re.search(regex, hs_search).group(2)
        print("contact ID is: " + result)
        simple_public_object_input = SimplePublicObjectInput(properties=properties)
        hs_client.crm.contacts.basic_api.update(contact_id=result, simple_public_object_input=simple_public_object_input)
        print("Contact has been updated and assigned to Porsha")

    elif bigtrip == "y" or "test":
        properties = {
            "lifecyclestage": "marketingqualifiedlead",
            "n2023_account_status": "Customer",
            "hubspot_owner_id": Amy_id
        }
        regex = r"('id': )'([\d]*)"
        public_object_search_request = PublicObjectSearchRequest(filter_groups=[{"filters":[{"value":requester,"propertyName":"email","operator":"EQ"}]}])
        hs_search = str(hs_client.crm.contacts.search_api.do_search(public_object_search_request=public_object_search_request))
        result = re.search(regex, hs_search).group(2)
        print("contact ID is: " + result)
        simple_public_object_input = SimplePublicObjectInput(properties=properties)
        hs_client.crm.contacts.basic_api.update(contact_id=result, simple_public_object_input=simple_public_object_input)
        print("Contact has been updated and assigned to Amy")

except ApiException as e:
    print("Exception when calling Hubspot API: %s\n" % e)

#Setting Email Parameters
msg['From'] = "***REMOVED***"
msg['To'] = ", ".join(recipients)
msg['Subject'] = f"Performance and Travel Form: Lead #{step} from {requester}"

#reading the html email from the external file.
data = open("email.html", "r").read()

#Email Body Content
hubspot_contact_link = f"https://app.hubspot.com/contacts/3057073/contact/{result}"
sheet_link = f"https://docs.google.com/spreadsheets/d/137HKP532tC3Y5zLl2igEenJl5IQChWVduow7Shh8ANk/edit#gid=1008713311&range=A{step}"
message = data.format(html_table = html_table, step = step, sheet_link = sheet_link, hubspot_contact_link = hubspot_contact_link)

#Add Message To Email Body
msg.attach(MIMEText(message, 'html'))

#To Send the Email
s.send_message(msg)

#Terminating the SMTP Session
s.quit()


#------------------
from gspread_formatting import *

# Sort the abridged Data frame from line 28 and copies it into a different sheet at a specified location and without headers or indices.
#append = int(input(f"""What row would you  like to append the Data to in '{ws2}'? """))
append = step
sorter = values.unstack().to_frame().T
set_with_dataframe(ws, sorter, row=append, col=3, include_index=False, include_column_header=False)

# Colors the row the appropriate color
if bigtrip == "n":
    gray = cellFormat(backgroundColor=color(0.7176470588235294,0.7176470588235294,0.7176470588235294))
    Porsha_yellow = cellFormat(backgroundColor=color(0.9450980392156862,0.7607843137254902,0.19607843137254902))
    format_cell_range(ws, f"B{append}:AVU{append}", gray)
    ws.update(f"A{step}", "Porsha")
    format_cell_range(ws, f"A{step}", Porsha_yellow)
elif bigtrip == "y" or "test":
    gray = cellFormat(backgroundColor=color(0.7176470588235294,0.7176470588235294,0.7176470588235294))
    Amy_blue = cellFormat(backgroundColor=color(0.43529411764705883,0.6588235294117647,0.8627450980392157))
    format_cell_range(ws, f"B{append}:AVU{append}", gray)
    ws.update(f"A{step}", "Amy")
    format_cell_range(ws, f"A{step}", Amy_blue)

# Move a cell from one sheet to another with A1 Notation.
mover = ws_live.acell('B2').value
ws.update(f'B{append}', mover)
set_row_height(ws, f'{step}', 21)

# Deletes the copied row mentioned in line 87 from ws1
ws_live.delete_rows(2, 2)
    
if bigtrip == "test":
    pass
else:
    append += 1
    stepper = open("step.py", "w")
    stepper.write(f"step = {append}")
    stepper.close()


