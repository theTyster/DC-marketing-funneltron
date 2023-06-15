#! python3

# followed the guide at www.thepythoncode.com/article/reading-emails-in-python




# cleantext for creating a folder this line is a function that creates folders without spaces and special characters.
#def clean(text):
#        return "".join(c if c.isalnum() else "_" for c in text)

def read():


    import imaplib
    import email
    from email.header import decode_header
    import re
    from auth.py import username, password

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
                        #print(body)
                else:
                    #extract content type of the email              
                    content_type = msg.get_content_type()           
                    if content_type == "text/plain":                
                        pass
                    #print(body)                                
    
    
    imap.close()
    imap.logout()
    
    
    match = re.search('WORKING', str(body), re.IGNORECASE)
    
    
    
    if match:
        print('itfreakinworked.jpg')
    
    else:
        print("uuuhhhhhh")
    
    '''
    this is not going to work because it will check the same emails over and over and over again. I need a way to move emails after they have been read to an archive.
    '''
    
    
read()
