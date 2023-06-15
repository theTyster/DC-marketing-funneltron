from oauth2client.service_account import ServiceAccountCredentials
import gspread


feed_auth = 'https://spreadsheets.google.com/feeds'
spreadsheets_auth = 'https://www.googleapis.com/auth/spreadsheets'
file_auth = 'https://www.googleapis.com/auth/drive.file' 
drive_auth = 'https://www.googleapis.com/auth/drive'  
json_loc = '/Users/tydavis/Documents/code/Work/Python/gsheet-cred.json'

scope = [feed_auth, spreadsheets_auth, file_auth, drive_auth]
creds = ServiceAccountCredentials.from_json_keyfile_name(json_loc, scope)
client = gspread.authorize(creds)


username = "***REMOVED***"
password = "***REMOVED***"