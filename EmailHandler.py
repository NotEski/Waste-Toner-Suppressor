import imaplib
import email
from email.header import decode_header
import os
import WasteTonerReporter
from datetime import datetime, timedelta

# all imports
# datetime, email, os, imaplib, json


def dateCleaner(date):
    "Tue, 9 Nov 2021 00:58:49 +0000"
    months = {"1":"Jan", "2":"Feb", "3":"Mar", "4":"Apr", "5":"May", "6":"Jun", "7":"Jul", "8":"Aug", "9":"Sep", "10":"Oct", "11":"Nov", "12":"Dec"}
    dateList = date.split(" ")[1:4]
    for i in months:
        if months[str(i)] == dateList[1]:
            dateList[1] = i
    return int(dateList[2] + dateList[1] + dateList[0].zfill(2))


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

def ReadEmail(username: str = "", password: str = "", server: str = "", port: int = 0, fromDaysAgo: int = 0, emailFolder: str = '"INBOX"') -> dict:
    print ("Connecting...", end="\r")
    imap = imaplib.IMAP4_SSL(str(server), int(port))
    
    # start timer thread

    # authenticate
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
    print (" " * len(f"\r{NoOfEmails} - {iteration} / {messages} : {iteration - NoOfEmails} / {UpperLimit}"), end="")
    print (": Emails checked", end="")
    print (f"\r{NoOfEmails} - {iteration} / {messages} : {iteration - NoOfEmails} / {UpperLimit}", end="\r")
    for i in range(messages, messages-messages, -1):
        emailDict = {}
        res, msg = imap.fetch(str(i), "(RFC822)") # fetch the email message by ID
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1]) # parse a bytes email into a message object
                subject, encoding = decode_header(msg["Subject"])[0] # decode the email subject
                if isinstance(subject, bytes):
                    subject = subject.decode()# if it's a bytes, decode to str
                From, encoding = decode_header(msg.get("From"))[0]# decode email sender
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                Date, encoding = decode_header(msg.get("Date"))[0]# decode email date
                if isinstance(Date, bytes):
                    Date = Date.decode(encoding)

                emailDict["subject"] = subject
                emailDict["from"] = From
                emailDict["date"] = dateCleaner(Date)
                if emailDict["date"] > 1000000:
                    emailDict["date"] = int (str (emailDict["date"])[:4] + "0" + str (emailDict["date"])[4:])
                    #print (emailDict["date"])
                if not fromDaysAgo == 0:
                    #print (str(emailDict["date"]) + " < " + str (int(((datetime.now() - timedelta(fromDaysAgo)).strftime("%Y%m%d")))))
                    if emailDict["date"] < (int(((datetime.now() - timedelta(fromDaysAgo)).strftime("%Y%m%d")))):
                        break
                if msg.is_multipart(): # if the email message is multipart
                    for part in msg.walk(): # iterate over email parts
                        content_type = part.get_content_type() # extract content type of email
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            body = (str(msg).split("<b=\nr>")[1].split("<br><br>")[0]) # get the email body
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition: # print text/plain emails and skip attachments
                            emailDict["body"] = body
                        elif "attachment" in content_disposition: # download attachment
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir("EmailDownloads\\" + folder_name): # make a folder for this email (named after the subject)
                                    os.mkdir("EmailDownloads\\" + folder_name)
                                filepath = os.path.join(folder_name, filename) # download attachment and save it
                                open("EmailDownloads\\" + filepath, "wb").write(part.get_payload(decode=True))
                                emailDict["attachment"] = filepath
                        elif "attachment" not in content_disposition:
                            emailDict["attachment"] = None
                else: # extract content type of email
                    content_type = msg.get_content_type() # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain": # print only text email parts
                        emailDict["body"] = body
                emailList.append(emailDict)
                NoOfEmails += 1
        iteration +=1
        print (" " * len(f"\r{NoOfEmails} - {iteration} / {messages} : {iteration - NoOfEmails} / {UpperLimit}"), end="")
        print (": Emails checked", end="")
        print (f"\r{NoOfEmails} -  {iteration} / {messages} : {iteration - NoOfEmails} / {UpperLimit}", end="\r")
        if (iteration - NoOfEmails) > UpperLimit:
            break
    imap.close()
    imap.logout()
    print ()
    #end timer thread and print time elapsed
    return emailList


if __name__ == "__main__":
    WasteTonerReporter.wasteToner()