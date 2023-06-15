#! python3

# followed the guide at www.thepythoncode.com/article/reading-emails-in-python


import imaplib
import email
from email.header import decode_header
import re


# cleantext for creating a folder this line is a function that creates folders without spaces and special characters.
def clean(text):
        return "".join(c if c.isalnum() else "_" for c in text)


#Creds
username = "***REMOVED***"
password = "***REMOVED***"

#create an IMAP4 class with SSL
imap = imaplib.IMAP4_SSL("imap.outlook.com")

#login
imap.login(username, password)

status, messages = imap.select("Inbox")


#total number of emails
messages = int(messages[0])

N = int(input("how many messages do you need to check?\n"))

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
    ok. from this point here is what I want to do:
    I need to create a script that will mark the designated line in the gsheets as being worked when someone responds that they are working on a project.
    this is going to be hard because the code will need to identify which line based on the context of the emails.
    Ideally this would be automated since I can not predict when my coworkers will respond (or if they will ever respond) wiht who is worknign on a specific deal.
    This will also require a bit of user education since in order for it to be marked people will need to know how to respond so that it will do that.
    it would be really cool if I could integrate this with todoist so that it would not only mark the gsheet but then also create a task for the individual who responded in Todoist and then when
    that task got completed it would mark the gsheet as "sent" so that people don't need ot keep coming back to their emails.

    ok. so here's what the experience looks like top to bottom:

    details are receved
    I start teh pythin code and send an email out to porsha, ***REMOVED***, and leanna.
    leanna responds in an email that she is going to work on it.
    code then marks the gsheet with "working" and creates a todoist and assigns it to Leanna.
    Leanna completes the todoist task and the code marks the gsheet as completed.

    this allows everyone to track who has been contacted and who has not from the marketing funnel.
    and I don't have to babysit anythng. XD

    ok so to make that work here are some things I need to do:

    I can set the input for the line to be a variable that gets printed into the email.

    when a respondant responds, then I could use regex to identify the keyword "working" and extract the three numbers next to it.
    this could then be set as a variable that the code could use to identify which line it needs to change the color of.

    That would also be important because if it is then going to be implemented into a todoist task that line on the Gsheets would be the easiest way to 
    extract the information and then compile it into another vertical list that is then made into a todoist task.

    Possibly the easiest way to do this would be to send the task as an email to the todoist board.
    would just reuse some old code and modify it so that it would perform the necessary task.

    up to this point I do not need to dip into the todoist python modules.
    holy crap its already almost midnight.
    lets take a sec and see how complicated those modules are.

    all I need to do is be able to trigger an event after someone marks their todoist task as complete.

    this would mean that I would need to have a client monitoring the state of a task. I don't have access ot the server-side code obviously so I cant set up a trigger after someone completes a task...
    right?
    maybe there is a way with a webhook? Could send an email? that would still require a cronjob to be monitoring incoming emails.
    I need something that can host some code and execute it when a certain trigger occurs from todoist.

    I feel like maybe a google app could do this? but that sounds kind of daunting. Maybe I should look into that a little bit though.
    is this why people use AWS?
    ookaaayyy I found a cool thing called python anywhere that I think will serve the purpose that I need.
    With this, I can possibly set up a cronjob. nope can't do that.

    it can potentially host the code and It looks like aI can set up a webapp that can probably monitor the state of other applications. maybe? I think that would probably use up too much CPU usage for the free version that I have.

