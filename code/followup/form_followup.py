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

## Function that determines who should be coppied on the resulting Email.

def censor():
    global bigtrip

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
msg['Subject'] = f"Performance and Travel Form: Lead #{step} from {requester}"

#Email Body Content
message = f"""
<h1>The Google Form "Performance and Travel" received a new submission!</h1><br>
<p>Click here to view the <a href="https://docs.google.com/forms/d/1b3wDerCr7vUCwbQ3sIulGc44_wYqAVJDLgha3ifSwhA/edit"><b>Google Form</b></a> this came from.</p>
<p>Click here to view the <a href="https://docs.google.com/spreadsheets/d/137HKP532tC3Y5zLl2igEenJl5IQChWVduow7Shh8ANk/edit?usp=sharing">spreadsheet data.</a></p>
<br>
<br>
--------------------START FORM RESPONSE--------------------
<br>
<p>&nbsp;</p>
{html_table}
<br>
----------------------END FORM RESPONSE--------------------
<br>
<br>
<h2>Meet Corey the Python</h2><br>
<img src="https://iili.io/jNNaDl.png" style="width: 400px;"><br>
<p>Corey is here to help you keep your tasks organized.</p>
<h4> If you are working on this registration update the status of it by "Replying to All" with:</h4>
<ul>
    <li>Building #{step}</li>
    <li>Sent #{step}</li>
    <li>Signed #{step}</li>
    <li>Lost #{step}</li>
    <li>Contacted #{step}</li>
</ul>
<p>Corey will then update the Google Sheet with the correct color. You are welcome to view the <a href="https://docs.google.com/spreadsheets/d/137HKP532tC3Y5zLl2igEenJl5IQChWVduow7Shh8ANk/edit?usp=sharing">spreadsheet</a> anytime to view your overall progress with these tasks.</p>
<p>A few things to note about corey:</p>
<ul>
    <li>Responses do not need to be case-sensitive. They do need to be spelled and spaced correctly and include the '#'.</li>
    <li>Responses do not need to be to this thread. You can email Corey from any email address or CC him on any thread and he will mark the correct line.</li>
    <li>To reset the celor back to gray for "Needs a Quote" just reply "reset #{step}".</li>
</ul>
<br>
<p>Contact <a href="mailto:***REMOVED***">Ty</a> if you have any suggestions or questions.</p>
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
#append = int(input(f"""What row would you  like to append the Data to in '{ws2}'? """))
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
