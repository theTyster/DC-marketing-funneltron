#! python3
# followed the guide at https://www.thepythoncode.com/article/deleting-emails-in-python and adapted for archiving.
#This file moves the first message from the inbox to the archive. so that Cell_label can read the next email.
def archiver():
    import imaplib
    from auth import username, password
    import email
    from email.header import decode_header

    #create an IMAP4 class with SSL
    imap = imaplib.IMAP4_SSL("imap.outlook.com")

    #login
    imap.login(username, password)
    imap.select("Inbox")

    """
    messages = messages[0].split(b' ')

    # This loop is merely for printing the subjects of the emails being archived
    for mail in messages:
    """

    status, messages = imap.select("Inbox")

    messages = int(messages[0])

    N = 1

    for i in range(messages, messages-N, -1):
        _, msg = imap.fetch(str(i), "(RFC822)")

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
        imap.copy(str(i), "Archive")
    for i in range(messages, messages-N, -1):
        _, msg = imap.fetch(str(i), "(RFC822)")

        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                #decode the email subject
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    #if it's a bytes type, decode to str
                    subject = subject.decode()
                print("removing "+subject+" from Inbox")
        #mark the mail as deleted
        imap.store(str(i), "+FLAGS", "\\Deleted")
    imap.expunge()
    imap.close()
    imap.logout()
