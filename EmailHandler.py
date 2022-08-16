import imaplib
import email
from email.header import decode_header
import os
import WasteTonerReporter
from datetime import datetime, timedelta



def dateCleaner(date):          # this function cleans the data format that gets passed in, from the email
    "Tue, 9 Nov 2021 00:58:49 +0000"
    months = {"1":"Jan", "2":"Feb", "3":"Mar", "4":"Apr", "5":"May", "6":"Jun", "7":"Jul", "8":"Aug", "9":"Sep", "10":"Oct", "11":"Nov", "12":"Dec"}
    dateList = date.split(" ")[1:4]
    for i in months:
        if months[str(i)] == dateList[1]:
            dateList[1] = i
    return int(dateList[2] + dateList[1] + dateList[0].zfill(2))    # returns the clean date


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

def ReadEmail(username: str = "", password: str = "", server: str = "", port: int = 0, fromDaysAgo: int = 0, emailFolder: str = '"INBOX"') -> dict:
    print ("Connecting...", end="\r")
    imap = imaplib.IMAP4_SSL(str(server), int(port))

    # Authenticate
    imap.login(username, password)


    status, messages = imap.select(emailFolder)
    # number of top emails to fetch
    # total number of emails
    messages = int(messages[0])


    UpperLimit = 20*fromDaysAgo
    
    emailList = []
    NoOfEmails = 0
    iteration = 0
    print ("Connected     ")
    print (" " * len(f"\r{NoOfEmails} - {iteration} / {messages} : {iteration - NoOfEmails} / {UpperLimit}"), end="")       # Prints a readable version of how far along it is in reading the emails  
    print (": Emails checked", end="")
    print (f"\r{NoOfEmails} - {iteration} / {messages} : {iteration - NoOfEmails} / {UpperLimit}", end="\r")
    for i in range(messages, messages-messages, -1):
        emailDict = {}
        res, msg = imap.fetch(str(i), "(RFC822)")                       # Fetch the email message by ID
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])             # Parse a bytes email into a message object
                subject, encoding = decode_header(msg["Subject"])[0]    # Decode the email subject
                if isinstance(subject, bytes):
                    subject = subject.decode()                          # If it's a bytes, decode to str
                From, encoding = decode_header(msg.get("From"))[0]      # Decode email sender
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                Date, encoding = decode_header(msg.get("Date"))[0]      # Decode email date
                if isinstance(Date, bytes):
                    Date = Date.decode(encoding)

                emailDict["subject"] = subject
                emailDict["from"] = From
                emailDict["date"] = dateCleaner(Date)
                if emailDict["date"] > 1000000:
                    emailDict["date"] = int (str (emailDict["date"])[:4] + "0" + str (emailDict["date"])[4:])
                if not fromDaysAgo == 0:
                    if emailDict["date"] < (int(((datetime.now() - timedelta(fromDaysAgo)).strftime("%Y%m%d")))):   # checks if its from the past x many days
                        break
                        
                        
                if msg.is_multipart():                              # If the email message is multipart
                    for part in msg.walk():                         # Iterate over email parts
                        content_type = part.get_content_type()      # Extract content type of email
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            body = (str(msg).split("<b=\nr>")[1].split("<br><br>")[0])  # Get the email body
                        except:
                            pass
                        emailDict["body"] = body
                            
                else:                                               # Extract content type of email
                    content_type = msg.get_content_type()           # Get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":                # Print only text email parts
                        emailDict["body"] = body
                emailList.append(emailDict)     # Add the email to the list of emails
                NoOfEmails += 1                 # Count up 1 to keep track of the amount of emails its gone through
        iteration +=1
        countString = f"\r{NoOfEmails} - {iteration} / {messages} : {iteration - NoOfEmails} / {UpperLimit}"
        print (" " * len(countString), end="")
        print (": Emails checked", end="")
        print (countString, end="\r")
        if (iteration - NoOfEmails) > UpperLimit:   # Does a check to see if more emails are older then x days old 
            break
    imap.close()
    imap.logout()   # Logs out of mailbox
    print ()
    return emailList


if __name__ == "__main__":
    WasteTonerReporter.wasteToner()
