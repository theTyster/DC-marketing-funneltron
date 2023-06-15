
import gspread
from auth import *

ws1 = 'Form Responses (Do not Edit)'
ws2 = "Porsha's Leads"
ws3 = "Amy's Leads"

ws_live = client.open('Performance and Travel Form (Responses)').worksheet(ws1)
wsp = client.open('Performance and Travel Form (Responses)').worksheet(ws2)
wsa = client.open('Performance and Travel Form (Responses)').worksheet(ws3)
wsp.update('C150', ' ')
wsp.update('C150', '')
wsa.update('C150', ' ')
wsa.update('C150', '')
