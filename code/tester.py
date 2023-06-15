#! python3

import gspread
from auth import *
from gspread_formatting import *


#Naming and gaining access to a specific workbook and specific sheets within that workbook.
ws1 = 'Form Responses (Do not Edit)'
ws2 = 'Form Responses (Current)'

ws_live = client.open('Performance and Travel Form (Responses)').worksheet(ws1)
ws = client.open('Performance and Travel Form (Responses)').worksheet(ws2)

# Colors the row the appropriate color
red = cellFormat(backgroundColor=color(255,0,0))
orange = cellFormat(backgroundColor=color(246,178,107))
yellow = cellFormat(backgroundColor=color(255,229,103))
green = cellFormat(backgroundColor=color(147,196,125))
purple = cellFormat(backgroundColor=color(194,123,160))
gray = cellFormat(backgroundColor=color(100, 100, 100))


append = 40

format_cell_range(ws, f"A{append}:AVU{append}", gray)
