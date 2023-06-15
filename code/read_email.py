#! python3
#This file provides the function that opens the first email in ***REMOVED***'s inbox and then returns body as a variable.#It is used by cell_label.py

# followed the guide at www.thepythoncode.com/article/reading-emails-in-python

# cleantext for creating a folder this line is a function that creates folders without spaces and special characters.
#def clean(text):
#        return "".join(c if c.isalnum() else "_" for c in text)
body = None
plain_body = None
inbox_empty = None
def read():

    import imaplib
    import email
    from email.header import decode_header
    import re
    from auth import username, password

    #Creds
    
    #create an IMAP4 class with SSL
    imap = imaplib.IMAP4_SSL("imap.outlook.com")
    
    #login
    imap.login(username, password)
    
    status, messages = imap.select("Inbox")
    
    
    #total number of emails
    messages = int(messages[0])
    
    N = 1
    
    for i in range(messages, messages-N, -1):
         try:
             res, msg = imap.fetch(str(i), "(RFC822)")
             imap.fetch
             for response in msg:
                 if isinstance(response, tuple):
                     #parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])
                    if msg.is_multipart():
                        # iterate over email parts
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            try:
                                global body
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                pass
                    else:
                        #extract content type of the email              
                        content_type = msg.get_content_type()           
                        if content_type == "text/plain":                
                            global plain_body
                            plain_body = msg.get_payload(decode=True).decode()
         except imaplib.IMAP4.error:
             inbox_empty = True
             return

    imap.close()
    imap.logout()
    
read()
try:
    print(body)
except:
    print(plain_body)
