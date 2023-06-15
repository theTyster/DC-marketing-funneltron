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
ws2 = 'Form Responses (Current)'

ws_live = client.open('Performance and Travel Form (Responses)').worksheet(ws1)
ws = client.open('Performance and Travel Form (Responses)').worksheet(ws2)


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

#Keywords to be searched in emails
building = 'BUILDING'
sent = 'SENT'
signed = 'SIGNED'
lost = 'LOST'
contacted = 'CONTACTED'
reset = 'RESET'


#Removes any quoted text from the body of the email.
#I'm a bit proud of this because this is my first time writing a class. this was written so that either plaintext or HTML emails would be considered.
class unquote:
    def __init__(self, bod):
        self.bod = bod
    def html(msg):
        try:
            regex = "<blockquote[\s\S]*"
            replace = re.sub(regex, '', msg.bod)
        except:
            regex = "Forwarded Message[\s\S*"
            replace = re.sub(regex, '', replace)
        return replace
    def plain(msg):
        try:
            regex = "^>.*$"
            replace = re.sub(regex, '', msg.bod, flags=re.MULTILINE)
        except:
            regex = "Forwarded Message[\s\S*"
            replace = re.sub(regex, '', replace)
        return replace


#Marks the designated line on the gsheet the designated color.
def paint(line, color):
    format_cell_range(ws, f"B{line}:AVU{line}", color)

'''
Considers whether anyone of the keywords listed above are in the email. 
If there exists one of the keywords it finds the designated line Number next to that keyword. 
It then passes the line number on to paint() and designates the color it should be painted.
'''
def highlighter(body_type):
    if re.search(building, str(body_type), re.IGNORECASE):
        line = re.search(building + ' ' + '#(\d\d?\d?)', body_type, re.IGNORECASE).group(1)
        paint(line, yellow)
        print(f"Marking line #{line} as 'building'")
    else: 
        pass
        
    if re.search(sent, str(body_type), re.IGNORECASE):
        line = re.search(sent + ' ' + '#(\d\d?\d?)', body_type, re.IGNORECASE).group(1)
        paint(line, orange)
        print(f"Marking line #{line} as 'sent'")
    else: 
        pass

    if re.search(signed, str(body_type), re.IGNORECASE):
        line = re.search(signed + ' ' + '#(\d\d?\d?)', body_type, re.IGNORECASE).group(1)
        paint(line, green)
        print(f"Marking line #{line} as 'signed'")
    else: 
        pass

    if re.search(lost, str(body_type), re.IGNORECASE):
        line = re.search(lost + ' ' + '#(\d\d?\d?)', body_type, re.IGNORECASE).group(1)
        paint(line, red)
        print(f"Marking line #{line} as 'lost'")
    else: 
        pass
        
    if re.search(contacted, str(body_type), re.IGNORECASE):
        line = re.search(contacted + ' ' + '#(\d\d?\d?)', body_type, re.IGNORECASE).group(1)
        paint(line, purple)
        print(f"Marking line #{line} as 'contacted'")
    else: 
        pass

    if re.search(reset, str(body_type), re.IGNORECASE):
        line = re.search(reset + ' ' + '#(\d\d?\d?)', body_type, re.IGNORECASE).group(1)
        paint(line, gray)
        print(f"Marking line #{line} as 'reset'")
    else: 
        pass

#pulls the body from the first email in the inbox
read_email.read()

'''
Checks whether the body is in HTML or Plaintext format.
Then, shortens the email by removing all quoted text.
highlights the gsheet.
'''
if read_email.body is not None:
    body = read_email.body
    print(body)
    shortened = unquote(body)
    html = shortened.html()
    print("HTML Email Detected")
    highlighter(html)

else:
    plain_body = read_email.plain_body
    shortened = unquote(plain_body)
    plain = shortened.plain()
    print("Plaintext Email Detected")
    highlighter(plain)
