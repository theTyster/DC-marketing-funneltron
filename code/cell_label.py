#! python3
#This file searches the body of an email for given keywords and then marks the google sheet the appropriate color on the appropriate line.
#In order for this file to execute correctly email responses need to be in the following format:
# <Email Status> #<Gsheet Row number>

import gspread
from auth import *
from gspread_formatting import *
import read_email
import re

#Naming and gaining access to a specific workbook and specific sheets within that workbook.
ws1 = 'Form Responses (Do not Edit)'
ws2 = "Porsha's Leads"
ws3 = "Amy's Leads"

ws_live = client.open('Performance and Travel Form (Responses)').worksheet(ws1)
wsp = client.open('Performance and Travel Form (Responses)').worksheet(ws2)
wsa = client.open('Performance and Travel Form (Responses)').worksheet(ws3)

#Colors to change the Gsheet rows
"""
Remember: for some reason this module needs RGB to be in the range of 0-1
even though true RGB is in the range of 0-255. To get the range that this 
code can use, you need to divide the true RGB values by 255.
"""
red = cellFormat(backgroundColor=color(1, 0, 0))
orange = cellFormat(backgroundColor=color(0.9647058823529412,0.6980392156862745,0.4196078431372549))
yellow = cellFormat(backgroundColor=color(1,0.8980392156862745,.6))
green = cellFormat(backgroundColor=color(0.5764705882352941,0.7686274509803922,0.49019607843137253))
purple = cellFormat(backgroundColor=color(0.7607843137254902,0.4823529411764706,0.6274509803921569))
gray = cellFormat(backgroundColor=color(0.7176470588235294,0.7176470588235294,0.7176470588235294))


def error_msg(message):

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


    #Setting Email Parameters
    msg['From'] = "***REMOVED***"
    msg['To'] = "***REMOVED***"
    msg['Subject'] = "Houston, we have a problem"

    #reading the html email from the external file.
    data = open("email.html", "r").read()

    #Add Message To Email Body
    msg.attach(MIMEText(message, 'html'))

    #To Send the Email
    s.send_message(msg)

    #Terminating the SMTP Session
    s.quit()

#Removes any quoted text from the body of the email.
#I'm a bit proud of this because this is my first time writing a class. this was written so that either plaintext or HTML emails would be considered.
class unquote:
    def __init__(self, bod):
        self.bod = bod
    def html(msg):
        try:
            regex = "<blockquote[\s\S]*"
            replace = re.sub(regex, '', msg.bod)
        except Exception:
            regex = "Forwarded Message[\s\S*"
            replace = re.sub(regex, '', replace)
        return replace
    def plain(msg):
        try:
            regex = "^>.*$"
            replace = re.sub(regex, '', msg.bod, flags=re.MULTILINE)
        except Exception:
            regex = "Forwarded Message[\s\S*"
            replace = re.sub(regex, '', replace)
        return replace


#Marks the designated line on the gsheet the designated color.
#Sends an email if there is an error.
def paint(line, status, color, amy_true, body, regex):
    if amy_true is True:
        format_cell_range(wsa, f"A{line}:AVU{line}", color)
        print(f"Marking line #{line} as {status}")
    elif amy_true is False:
        format_cell_range(wsp, f"A{line}:AVU{line}", color)
        print(f"Marking line #{line} as {status}")
    else:
        amy_true = str(amy_true)
        error = f"""
<h3>There was an error reading the contents of the message in Corey's inbox.</h3>
<br>
<p>Amy's match came back as: {amy_true}</p>
<br>
<h3>Here is the current regex in use:</h3>
    <p>{regex}</p>
<br>
<h3>This is the contents of the body of the email that caused the error:</h3>

    <p>{body}</p>
"""
        error_msg(error)
        print("There was an error reading the contents of the message. An email notification has been sent.")
        exit()


'''
Considers whether anyone of the keywords listed above are in the email. 
If there exists one of the keywords it finds the designated line Number next to that keyword. 
It then passes the line number on to paint() and designates the color it should be painted.
'''
def highlighter(body_type):

    #Keywords to be searched in emails
    building = 'BUILDING #'
    sent = 'SENT #'
    signed = 'SIGNED #'
    lost = 'LOST #'
    contacted = 'CONTACTED #'
    reset = 'RESET #'

    if re.search(building, str(body_type), re.IGNORECASE):
        building = 'BUILDING'
        what_row_regex = r'#(\d\d?\d?)'
        amy_regex = r"***REMOVED***"
        regex ='<br>' + what_row_regex + '<br>' + amy_regex
        line = re.search(building + ' ' + what_row_regex, body_type, re.IGNORECASE).group(1)
        amy_consultant = bool(re.search(amy_regex, body_type, re.IGNORECASE))
        paint(line, building, yellow, amy_consultant, body_type, regex)
    else: 
        pass

    if re.search(sent, str(body_type), re.IGNORECASE):
        sent = 'SENT'
        what_row_regex = r'#(\d\d?\d?)'
        amy_regex = r"***REMOVED***"
        regex ='<br>' + what_row_regex + '<br>' + amy_regex
        line = re.search(sent + ' ' + what_row_regex, body_type, re.IGNORECASE).group(1)
        #line = re.compile(body_type, re.I).match(sent + ' ' + what_row_regex).group(1)
        amy_consultant = bool(re.search(amy_regex, body_type, re.IGNORECASE))
        paint(line, sent, orange, amy_consultant, body_type, regex)
    else: 
        pass

    if re.search(signed, str(body_type), re.IGNORECASE):
        signed = 'SIGNED'
        what_row_regex =r'#(\d\d?\d?)'
        amy_regex = r"***REMOVED***"
        regex ='<br>' + what_row_regex + '<br>' + amy_regex
        line = re.search(signed + ' ' + what_row_regex, body_type, re.IGNORECASE).group(1)
        amy_consultant = bool(re.search(amy_regex, body_type, re.IGNORECASE))
        paint(line, signed, green, amy_consultant, body_type, regex)
    else: 
        pass

    if re.search(lost, str(body_type), re.IGNORECASE):
        lost = 'LOST'
        what_row_regex =r'#(\d\d?\d?)'
        amy_regex = r"***REMOVED***"
        regex ='<br>' + what_row_regex + '<br>' + amy_regex
        line = re.search(lost + ' ' + what_row_regex, body_type, re.IGNORECASE).group(1)
        amy_consultant = bool(re.search(amy_regex, body_type, re.IGNORECASE))
        paint(line, lost, red, amy_consultant, body_type, regex)
    else: 
        pass

    if re.search(contacted, str(body_type), re.IGNORECASE):
        contacted = 'CONTACTED'
        what_row_regex =r'#(\d\d?\d?)'
        amy_regex = r"***REMOVED***"
        regex ='<br>' + what_row_regex + '<br>' + amy_regex
        line = re.search(contacted + ' ' + what_row_regex, body_type, re.IGNORECASE).group(1)
        amy_consultant = bool(re.search(amy_regex, body_type, re.IGNORECASE))
        paint(line, contacted, purple, amy_consultant, body_type, regex)
    else: 
        pass

    if re.search(reset, str(body_type), re.IGNORECASE):
        reset = 'RESET'
        what_row_regex =r'#(\d\d?\d?)'
        amy_regex = r"***REMOVED***"
        regex ='<br>' + what_row_regex + '<br>' + amy_regex
        line = re.search(reset + ' ' + what_row_regex, body_type, re.IGNORECASE).group(1)
        amy_consultant = bool(re.search(amy_regex, body_type, re.IGNORECASE))
        paint(line, reset, gray, amy_consultant, body_type, regex)
    else: 
        pass

#pulls the body from the first email in the inbox

'''
Checks whether the body is in HTML or Plaintext format.
Then, shortens the email by removing all quoted text.
highlights the gsheet.
'''

def loop_plain():
    read_email.read()
    if read_email.inbox_empty == True:
        print("There are no new messages in your inbox") 
        return None
    else:
        plain_body = read_email.plain_body
        shortened = unquote(plain_body)
        plain = shortened.plain()
        print("Plaintext Email Detected")
        highlighter(plain)
        archive_email.archiver()
        loop_plain()

def loop_html():
    read_email.read()
    if read_email.inbox_empty == True:
        print("There are no new messages in your inbox") 
        return None
    else:
        body = read_email.body
        shortened = unquote(body)
        print("HTML Email Detected")
        html = shortened.html()
        highlighter(html)
        archive_email.archiver()
        loop_html()

import archive_email

try:    
    loop_plain()
except TypeError:
    print("No plaintext emails")
except UnboundLocalError: 
    pass

try:
    loop_html()
except TypeError: 
    print("No HTML emails")
except UnboundLocalError: 
    pass


#Changes the values of the bottom-most cells so that the color updater will work.
wsp.update('C150', ' ')
wsp.update('C150', '')
wsa.update('C150', ' ')
wsa.update('C150', '')
print("I tried")
