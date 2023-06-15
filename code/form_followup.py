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
ws2 = "Porsha's Leads"
ws3 = "Amy's Leads"

ws_live = client.open('Performance and Travel Form (Responses)').worksheet(ws1)
wsp = client.open('Performance and Travel Form (Responses)').worksheet(ws2)
wsa = client.open('Performance and Travel Form (Responses)').worksheet(ws3)

# assigns a row into a Panda data frame.
try:
    values = pd.DataFrame(ws_live.row_values(2))
except IndexError:
    pepe = """
                               .';:cccc:;,'.                   ....                        
                           .,ccccccccccccccccc;'.      ..,;cccccccccc;.                    
                         ':cccccccccccccccccccccc:...;ccccccccccccccccc:.                  
                       .ccccccccccccccccccccccccccc,ccccccccccccccccccccc'                 
                      ;cccccccccccc:::::::::::::ccc,cccccccccccccccccccccc;                
                    .:ccccccccc::::cccccccccccc::::':::::::::::::::::::::::.               
                   .cccccccc;;:cccccccccccccccccccc:;:cccccccccccccccccccccc;..            
                  'cccccccc:ccccccccccccccc::::::::::,:::ccccccccc:::::::::::::;'.         
                 ,cccccccccccccccccccc::::::::cccccccccc::;cc:::::c::::::::::::cc:;.       
               .:cccccccccccccccccc:;;:;:::::::ccccccccc::;;::::::::cccccccccc::::::,      
            .,:;cccccccccccccccc:::;;;:cccccccccccccccccccc:;ccccccccccc::::;;;;;:ccc;.    
          .;cc,ccccccccccccc:::::::cccccccccc;,..,lkkOOOOkkx,cc;:ldxkO0d.     .:KX0Oxl;'   
         ,cccc,cccccccccccc;;:ccccccc::;:ldc.'.     ,OMMMMMM;d0WMMMMMW; :,  ..   OMMMMN.   
        ;cccc:cccccccccccccccc:;';oO0KNMMW;  XK ,xd.  kMMMWoXMMMMMMMW.  '. dMM,   WMMNx.   
       ,cccccccccccccccccccccccc;c:;lx0WM'   o, OMMd  'W0d;;cx0KNMMMk   ';  ..    ll;.     
      .cccccccccccccccccccccccccc:::::c;:... ..  .....';;:;:ccc:;;:::'.....'',;ccc,        
      ccccccccccccccccccccccccccccccc:::::::::::::::::;,;cccccccccccccccccccccc:'.         
     :cccccccccccccccccccccccccccccccccccccccccccccc;;:ccccccc:::;::::::::::..             
    ;cccccccccccccccccccccccccccccccccccccccccc:::::ccccccccccccc:;:::cccccc'              
  .:ccccccccccccccccccccccccccccccccccccccccc::cccccccccccccccccccccccccccccc:.            
 .ccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc:.          
 cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc.         
.cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc:         
 ccccccccccccccccccccccccccccc;,,,,,,,,,,,,,;:cccccccccccccccccccccccccccccccccccc'        
 ,cccccccccccccccccccccccccc,,;::::::::::::::;,,,;;;;;;::ccccccccccccccc:::;;;;,,,::       
  ccccccccccccccccccccccccc.::::;;;;;;;;;;;;;;::::::::;;;,,,,,,,,,,,,,,,;;;::::::::.       
  .cccccccccccccccccccccccc.::::::::::::::::::;;;;;;;;:::::::::::::::::::::;;;;;.          
   ,cccccccccccccccccc,ccccc;,,,,,,,,,,,,,;::::::::::;;;;;;;;;;;;;;;;;;;;;;:::::           
    .:ccccccccccccccccc;:ccccccccccccccccc:;;;,,,,,,,,,,,;;;::::::::::::::::;'.            
      .;ccccccccccccccccc::cccccccccccccccccccccccccccccc::;;;;;;;;;;;;;;;.                
     ....';:ccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc;.                  
     '''''.',;:::::::cccccccccccccccccccccccccccccccccccccccccccccc;'.                     
     ''''''''.',,:cc::::::::::::ccccccccccccccccccccccccccccc:,..                          
     ''''''''''''...',,;:cccccc::::::::::::::::::::::::::::,'.                             
     '''''''''''''''''''...''',,;;:cccccccccccccccccc:;,,'..'''..                          
     '''''''''''''''''''''''''''''...''''',,,,'''''...'''''''''''''.                       
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''.                     
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
"""
    print(pepe + "\n" + "No new leads. Check back later, player.")
    exit()
requester = ws_live.acell('C2').value
leads_leftp = int(wsp.acell("F1").value)+1
leads_lefta = int(wsa.acell("F1").value)+1


# remove all empty values from data frame by replacing them with NaN and then dropping all cells with Nan value. inplace=True specifies that the cells will be replaced as opposed to appended I guess.
values[0].replace('', np.nan, inplace=True)
values.dropna(subset=[0], inplace=True)

# reindexes. So that "values" is only the raw data formatted as a list.
values = values.set_index([pd.Index(list(range(0, len(values.index))))])

#Prints the form data into std out to make it easier to determine who the data should go to.
#Asks what line the school is on for further organization.
print(values)

school_name_row = input("What line is the school name on? (default is line 11) \n")

if school_name_row == "":
    school_name_row = 11
    school = values.iat[school_name_row,0]
else:
    while int(school_name_row) > int(len(values.index)) or int(school_name_row) < 1: 
        print("Input is out of range\n")
        school_name_row = int(input("What line is the school name on? (default is line 11) \n"))

    school = str(values.iat[int(school_name_row),0])

    
values = values.append([0])
values = values.shift(periods=1)
tmp = values.iat[int(school_name_row)+1,0]
values.iat[int(school_name_row)+1,0] = values.iat[0,0]
values.iat[0,0] = tmp
values.dropna(subset=[0], inplace=True)

print("School is " + school + "\n")
print(values)



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


from step import *
# add a column header, hide the index, and export to HTML
if bigtrip == "n":
    values.rename(columns = {0:f"<p style='color: darkgreen;'>Performance and Travel Form Response #{stepp}</p>"}, inplace=True)
    html_table = values.style.hide_index().to_html()
elif bigtrip == "y" or "test":
    values.rename(columns = {0:f"<p style='color: darkgreen;'>Performance and Travel Form Response #{stepa}</p>"}, inplace=True)
    html_table = values.style.hide_index().to_html()

#------------------




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
        try:
            result = re.search(regex, hs_search).group(2)
            print("contact ID is: " + result)
            simple_public_object_input = SimplePublicObjectInput(properties=properties)
            hs_client.crm.contacts.basic_api.update(contact_id=result, simple_public_object_input=simple_public_object_input)
            print("Contact has been updated and assigned to Porsha")
        except AttributeError:
            print(f"""**********BZZZZZZZZZZZTT**********\nContact does not exist in Hubspot try searching for them at the link below: \n https://app.hubspot.com/contacts/3057073/objects/0-1/views/all/list?query={requester}\n**********THATS AN ERROR**********""")
            exit()
    elif bigtrip == "y":
        properties = {
            "lifecyclestage": "marketingqualifiedlead",
            "n2023_account_status": "Customer",
            "hubspot_owner_id": Amy_id
        }
        regex = r"('id': )'([\d]*)"
        public_object_search_request = PublicObjectSearchRequest(filter_groups=[{"filters":[{"value":requester,"propertyName":"email","operator":"EQ"}]}])
        hs_search = str(hs_client.crm.contacts.search_api.do_search(public_object_search_request=public_object_search_request))
        try:
            result = re.search(regex, hs_search).group(2)
            print("contact ID is: " + result)
            simple_public_object_input = SimplePublicObjectInput(properties=properties)
            hs_client.crm.contacts.basic_api.update(contact_id=result, simple_public_object_input=simple_public_object_input)
            print("Contact has been updated and assigned to Amy")
        except AttributeError:
            print(f"""**********BZZZZZZZZZZZTT**********\nContact does not exist in Hubspot try searching for them at the link below: \n https://app.hubspot.com/contacts/3057073/objects/0-1/views/all/list?query={requester}\n**********THATS AN ERROR**********""")
            exit()
    elif bigtrip == "test":
        result = "test"
        pass
except ApiException as e:
    print("Exception when calling Hubspot API: %s\n" % e)


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


from gspread_formatting import *

if bigtrip == "n":
    #Setting Email Parameters
    msg['From'] = "***REMOVED***"
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = f"Performance and Travel Form: Lead #{stepp} from {school}; {requester}"

    #reading the html email from the external file.
    data = open("email.html", "r").read()

    #Email Body Content
    hubspot_contact_link = f"https://app.hubspot.com/contacts/3057073/contact/{result}"
    sheet_link = f"https://docs.google.com/spreadsheets/d/137HKP532tC3Y5zLl2igEenJl5IQChWVduow7Shh8ANk/edit#gid=1279253480&range=A{stepp}"
    message = data.format(leads_left = leads_leftp, html_table = html_table, step = stepp, sheet_link = sheet_link, hubspot_contact_link = hubspot_contact_link, school = school)

    #Add Message To Email Body
    msg.attach(MIMEText(message, 'html'))

    #To Send the Email
    s.send_message(msg)

    #Terminating the SMTP Session
    s.quit()


    # Sort the abridged Data frame from line 28 and copies it into a different sheet at a specified location and without headers or indices.
    appendp = stepp
    appenda = stepa
    sorter = values.unstack().to_frame().T
    set_with_dataframe(wsp, sorter, row=appendp, col=2, include_index=False, include_column_header=False)

    # Colors the row the appropriate color
    gray = cellFormat(backgroundColor=color(0.7176470588235294,0.7176470588235294,0.7176470588235294))
    Porsha_yellow = cellFormat(backgroundColor=color(0.9450980392156862,0.7607843137254902,0.19607843137254902))
    format_cell_range(wsp, f"A{appendp}:AVU{appendp}", gray)
    # Move a cell from one sheet to another with A1 Notation.
    mover = ws_live.acell('B2').value
    wsp.update(f'A{appendp}', mover)
    mover = wsp.get(f'D{stepp}:AVU{stepp}')
    wsp.update(f'C{appendp}:AVU{appendp}', mover)
    set_row_height(wsp, f'{stepp}', 21)

    # Deletes the copied line from ws1
    ws_live.delete_rows(2)

elif bigtrip == "y" or "test":
    #Setting Email Parameters
    msg['From'] = "***REMOVED***"
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = f"Performance and Travel Form: Lead #{stepa} from {requester}"

    #reading the html email from the external file.
    data = open("email.html", "r").read()

    #Email Body Content
    hubspot_contact_link = f"https://app.hubspot.com/contacts/3057073/contact/{result}"
    sheet_link = f"https://docs.google.com/spreadsheets/d/137HKP532tC3Y5zLl2igEenJl5IQChWVduow7Shh8ANk/edit#gid=1519048498&range=A{stepa}"
    message = data.format(leads_left = leads_lefta, html_table = html_table, step = stepa, sheet_link = sheet_link, hubspot_contact_link = hubspot_contact_link, school = school)

    #Add Message To Email Body
    msg.attach(MIMEText(message, 'html'))

    #To Send the Email
    s.send_message(msg)

    #Terminating the SMTP Session
    s.quit()


    # Sort the abridged Data frame from line 28 and copies it into a different sheet at a specified location and without headers or indices.
    appendp = stepp
    appenda = stepa
    sorter = values.unstack().to_frame().T
    set_with_dataframe(wsa, sorter, row=appenda, col=2, include_index=False, include_column_header=False)
# Colors the row the appropriate color
    gray = cellFormat(backgroundColor=color(0.7176470588235294,0.7176470588235294,0.7176470588235294))
    Amy_blue = cellFormat(backgroundColor=color(0.43529411764705883,0.6588235294117647,0.8627450980392157))
    format_cell_range(wsa, f"A{appenda}:AVU{appenda}", gray)
# Move a cell from one sheet to another with A1 Notation.
    mover = ws_live.acell('B2').value
    wsa.update(f'A{appenda}', mover)
    mover = wsa.get(f'D{stepa}:AVU{stepa}')
    wsa.update(f'C{appenda}:AVU{appenda}', mover)
    set_row_height(wsa, f'{stepa}', 21)
# Deletes the copied line from ws1
    ws_live.delete_rows(2)




if bigtrip == "test":
    pass
if bigtrip == "y":
    appenda += 1
    stepper = open("step.py", "w")
    writethis = f"""
stepp = {appendp}
stepa = {appenda}
"""
    stepper.write(writethis)
    stepper.close()
elif bigtrip == "n":
    appendp += 1
    stepper = open("step.py", "w")
    writethis = f"""
stepp = {appendp}
stepa = {appenda}
"""
    stepper.write(writethis)
    stepper.close()
