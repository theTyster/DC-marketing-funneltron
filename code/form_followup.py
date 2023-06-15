#! python3

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

# remove all empty values from data frame by replacing them with NaN and then dropping all cells with Nan value. idk what inplace=True means.
values[0].replace('', np.nan, inplace=True)
values.dropna(subset=[0], inplace=True)

# drop indices. So that "values" is only the raw data formatted as a list.
values = values.drop(values.index[0])

html_table = values.style.hide_columns().hide_index().render()
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

## Function that determines who should be coppied on the resulting Email.

def censor():
    bigtrip = input("Should Jon be included on this? (y/N) \n")
    
    global recipients

    if bigtrip == "y":
    
        recipients = ["***REMOVED***", "***REMOVED***@***REMOVED***.com", "***REMOVED***"]
        print("Sending to Ty, Jon and Porsha")
    
    elif bigtrip == "n":
    
        recipients = ["***REMOVED***", "***REMOVED***"]
        print("Sending to Ty, Porsha")

    elif bigtrip == "":

        recipients = ["***REMOVED***", "***REMOVED***"]
        print("Sending to Ty and Porsha")
    
    elif bigtrip == "test":
        recipients =["***REMOVED***"]
        print("Sending you a test email")

    else:
        print("input is not recognized")
        censor()

censor()

#Setting Email Parameters
msg['From'] = "***REMOVED***"
msg['To'] = ", ".join(recipients)
msg['Subject'] = "Performance and Travel Form: New request for a proposal"

#Email Body Content
message = f"""
<h1>The Google Form "Performance and Travel" received a new submission!</h1><br>
<p>The details submitted are below.<br>
Hit "reply all" if you have questions regarding this email.<p>
<p>Click here to view the <a href="https://docs.google.com/forms/d/1b3wDerCr7vUCwbQ3sIulGc44_wYqAVJDLgha3ifSwhA/edit"><b>Google Form</b></a> this came from.</p>
<p>Click here to view the <a href="https://docs.google.com/spreadsheets/d/137HKP532tC3Y5zLl2igEenJl5IQChWVduow7Shh8ANk/edit?usp=sharing">spreadsheet data.</a></p>
<br>
<br>
<h4> If you are working on this registration update the status of it by "Replying to All" with:</h4>
<ul>
    <li>Working</li>
    <li>Sent</li>
</ul>
<p>This will help ensure that we do not have multiple people working on the same registration. Contact <a href="mailto:***REMOVED***">Ty</a> if you have any suggestions or questions.</p>
<br>
--------------------START FORM RESPONSE--------------------
<br>
<p>&nbsp;</p>
{html_table}
<br>
----------------------END FORM RESPONSE--------------------
"""
# Asks for a number and checks if a number was given through input.
#x=input("Which row would you like to send? ")
#try:
#    val = int(x)
#except ValueError:
#    print("That's not an int!")


#Add Message To Email Body
msg.attach(MIMEText(message, 'html'))

#To Send the Email
s.send_message(msg)

#Terminating the SMTP Session
s.quit()


#------------------
from gspread_formatting import *

# Sort the abridged Data frame from line 28 and copies it into a different sheet at a specified location and without headers or indices.
append = int(input(f"""What row would you  like to append the Data to in '{ws2}'? """))
sorter = values.unstack().to_frame().T
set_with_dataframe(ws, sorter, row=append, col=3, include_index=False, include_column_header=False)


# Colors the row the appropriate color
red = cellFormat(backgroundColor=color(255,0,0))
orange = cellFormat(backgroundColor=color(246,178,107))
yellow = cellFormat(backgroundColor=color(255,229,103))
green = cellFormat(backgroundColor=color(147,196,125))
purple = cellFormat(backgroundColor=color(194,123,160))
gray = cellFormat(backgroundColor=color(100, 100, 100))

format_cell_range(ws, f"A{append}:AVU{append}", gray)


# Move a cell from one sheet to another with A1 Notation.
mover = ws_live.acell('B2').value
ws.update(f'B{append}', mover)

# Move a row from one sheet to the other with in a certain range of columns
#mover = ws_live.row_values(2)
#ws.append_row(mover, table_range="B")

# Deletes the copied row mentioned in line 87 from ws1
ws_live.delete_rows(2, 2)
