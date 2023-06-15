#! python3

# followed the guide at https://www.thepythoncode.com/article/deleting-emails-in-python and adapted for archiving.

import imaplib
import email
from email.header import decode_header


#Creds
username = "***REMOVED***"
password = "***REMOVED***"

#create an IMAP4 class with SSL
imap = imaplib.IMAP4_SSL("imap.outlook.com")

#login
imap.login(username, password)
imap.select("Inbox")




status, messages = imap.search(None, "SEEN")

messages = messages[0].split(b' ')

# This loop is merely for printing the subjects of the emails being archived
for mail in messages:
    _, msg = imap.fetch(mail, "(RFC822)")

    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            #decode the email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                #if it's a bytes type, decode to str
                subject = subject.decode()
            print("Archiving", subject)
    #mark the mail as deleted
    imap.store(mail, "+FLAGS", "Archive")

imap.expunge()
imap.close()
imap.logout()



# get help here: https://github.com/awangga/outlook
# or here: https://stackoverflow.com/questions/122267/imap-how-to-move-a-message-from-one-folder-to-another
