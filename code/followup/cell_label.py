#! python3
#This file searches the body of an email for given keywords and then marks the google sheet the appropriate color on the appropriate line.

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

# Colors the row the appropriate color.
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

read_email.read()

#searches for atleast one digit (max of 3) preceded by '#' in the body of the first email in outlook
#.group(1) returns a string of the matched text. Setting the option as '1' returns the string in the first parenthesis.
#read more at this link https://docs.python.org/3/library/re.html#re.Match.group
line = re.search('#(\d\d?\d?)', str(read_email.body), re.IGNORECASE).group(1)

building = 'BUILDING'
sent = 'SENT'
signed = 'SIGNED'
lost = 'LOST'
contacted = 'CONTACTED'
reset = 'RESET'

def red_paint():
    format_cell_range(ws, f"A{line}:AVU{line}", red)
def orange_paint():
    format_cell_range(ws, f"A{line}:AVU{line}", orange)
def yellow_paint():
    format_cell_range(ws, f"A{line}:AVU{line}", yellow)
def green_paint():
    format_cell_range(ws, f"A{line}:AVU{line}", green)
def purple_paint():
    format_cell_range(ws, f"A{line}:AVU{line}", purple)
def gray_paint():
    format_cell_range(ws, f"A{line}:AVU{line}", gray)


if  re.search(building, str(read_email.body), re.IGNORECASE):

    yellow_paint() 
    print(f"Marking line #{line} as 'building'"
    
elif  re.search(sent, str(read_email.body), re.IGNORECASE):

    orange_paint() 
    print(f"Marking line #{line} as 'sent'"

elif  re.search(signed, str(read_email.body), re.IGNORECASE):

    green_paint()
    print(f"Marking line #{line} as 'signed'"

elif  re.search(lost, str(read_email.body), re.IGNORECASE):

    red_paint()
    print(f"Marking line #{line} as 'lost'"
    
elif  re.search(contacted, str(read_email.body), re.IGNORECASE):

    purple_paint()
    print(f"Marking line #{line} as 'contacted'"

elif  re.search(reset, str(read_email.body), re.IGNORECASE):

    gray_paint()
    print(f"Marking line #{line} as 'reset'"

else:
    pass
    
