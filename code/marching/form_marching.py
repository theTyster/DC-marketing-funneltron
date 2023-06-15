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

ws_live = client.open('Marching Championship Form (Responses)').worksheet(ws1)
ws = client.open('Marching Championship Form (Responses)').worksheet(ws2)

# assigns a row into a Panda data frame.
values = pd.DataFrame(ws_live.row_values(2))
requester = ws_live.acell('C2').value
print(requester)

# remove all empty values from data frame by replacing them with NaN and then dropping all cells with Nan value. inplace=True specifies that the cells will be replaced as opposed to appended I guess.
values[0].replace('', np.nan, inplace=True)
values.dropna(subset=[0], inplace=True)

# drop indices. So that "values" is only the raw data formatted as a list.
values = values.drop(values.index[0])

from step import *
# add a column header, hide the index, and export to HTML
values.rename(columns = {0:f"<p style='color: darkgreen;'>Marching Championship Form Response #{step}</p>"}, inplace=True)
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

## Function that determines who should be coppied on the resulting Email.

def censor():
    global bigtrip
    bigtrip = input("Should Alan be included on this? (y/N) \n")
    
    global recipients

    if bigtrip == "y":
    
        recipients = ["***REMOVED***", "alan.hanna@***REMOVED***.com", "***REMOVED***"]
        print("Sending to Ty, Alan and Amy")
    
    elif bigtrip == "n":
    
        recipients = ["***REMOVED***", "Amy.locke@***REMOVED***.com"]
        print("Sending to Ty, Amy")

    elif bigtrip == "":

        recipients = ["***REMOVED***", "Amy.locke@***REMOVED***.com"]
        print("Sending to Ty and Amy")
    
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
msg['Subject'] = f"Marching Championship Form: lead #{step} from {requester}"

#Email Body Content
message = f"""
<h1>The Google Form "Marching Championship" received a new submission!</h1><br>
<p>The details submitted are below.<br>
Hit "reply all" if you have questions regarding this email.<p>
<p>Click here to view the <a href="https://docs.google.com/forms/d/10ayQbaqGoztJGUbfH5utasTDQJtUw_Vuv_gmbhFU2gM/edit"><b>Google Form</b></a> this came from.</p>
<p>Click here to view the <a href="https://docs.google.com/spreadsheets/d/1c2vbZKf71JEyj55aVx05QGSZXMWPc2BIzNqFVBJ_zno/edit?resourcekey#gid=1653083831">spreadsheet data.</a></p>
<br>
<p>Note: Corey can't work on this sheet.<br> You will need to update the colors on the google sheet manually.</p>
<img src="https://iili.io/jNNlx2.png" style="width: 400px;">
<br>
--------------------START FORM RESPONSE--------------------
<br>
<p>&nbsp;</p>
{html_table}
<br>
<br>
---------------------END FORM RESPONSE---------------------
"""

#Add Message To Email Body
msg.attach(MIMEText(message, 'html'))

#To Send the Email
s.send_message(msg)

#Terminating the SMTP Session
s.quit()


#------------------
from gspread_formatting import *

# Sort the abridged Data frame from line 28 and copies it into a different sheet at a specified location and without headers or indices.
append = step
sorter = values.unstack().to_frame().T
set_with_dataframe(ws, sorter, row=append, col=3, include_index=False, include_column_header=False)

# Colors the row the appropriate color
gray = cellFormat(backgroundColor=color(0.7176470588235294,0.7176470588235294,0.7176470588235294))
format_cell_range(ws, f"A{append}:AVU{append}", gray)

# Move a cell from one sheet to another with A1 Notation.
mover = ws_live.acell('B2').value
ws.update(f'B{append}', mover)

# Deletes the copied row mentioned in line 87 from ws1
ws_live.delete_rows(2, 2)

if bigtrip == "test":
    pass
else:
    append += 1
    stepper = open("step.py", "w")
    stepper.write(f"step = {append}")
    stepper.close()
