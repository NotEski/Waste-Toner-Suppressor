import EmailHandler
import json, os
from datetime import datetime
from datetime import timedelta
import urllib.request

print ("Loading Files...", end="\r")

#pyinstaller --onefile --hidden-import os --hidden-import email --hidden-import datetime --hidden-import urllib.requests --hidden-import imaplib --hidden-import emailhandler.py WasteTonerReporter.py
defaultConfig = '{"email":"YOUR EMAIL","password":"YOUR EMAILS PASSWORD","server":"outlook.office365.com","port":993,"FromXDaysAgo":4,"KeepFromXDaysAgo":30,"emailFolder":"INBOX"}'


if "config.json" not in os.listdir():
    file = open ("config.json", "w")
    file.write(defaultConfig)
    file.close()
else:
    with open ("config.json", "r") as file:
        config = file.read()
        if "YOUR EMAIL" in config:
            print ("Please fill out the config file and enter your email address and set up SMTP details")
            input ()
            quit ()


try:
    urllib.request.urlopen('http://google.com')
except:
    print ("ERROR - No Internet Connection")
    print ("Press 'Enter' to exit the program")
    input ("")
    quit()

with open ("config.json", "r") as configFile:
    config = json.load(configFile)
    email = config["email"]
    password = config["password"]
    server=config["server"]
    port=config["port"]
    daysAgo=config["FromXDaysAgo"]
    dataBaseKeepFrom=config["KeepFromXDaysAgo"]
    inboxFolder= '"' + config["emailFolder"].upper() + '"'

def wasteToner():
    print ("Loaded           ")
    print ("Connecting...", end="\r")
    emails = EmailHandler.ReadEmail(username=email, password=password, server=server, port=port, fromDaysAgo=daysAgo, emailFolder=inboxFolder) # retreives alerts as emails
    SupressMe = {}
    alerts = {}
    for i in emails: # loops through all emails
        if "e-BRIDGE CloudConnect" not in i["subject"]:
            continue
        #items = i["body"].split("<br>")
        try:
            items = i["body"].split("<br>")
            if len(items) < 5:
                items = items[0:-2] + items[-1].split("<br=\n>")
            SerialNumber = items[0].replace("Serial Number: ", "")
            Customer = items[2].replace("Description: ", "").replace("\n", "")
            Address = items[3].replace("Address: ", "").replace("\n", "") + ", " + items[4].replace("\n", "")
            ID = "".join(filter(str.isdigit, items[5].replace("Machine ID: ", "").replace("\n", "")))
            alerts[SerialNumber] = {"Customer": Customer, "Address": Address, "ID":ID}
        except Exception as e:
            pass
            #print (items)

    # check if the serial number is in the dictionarys for the past 30 days (as soon as it is found add it to the new SupressMe and start on the next one)
    # if not found add to suppress me and add to send a new alert
    # 
    # fetch an email check if its from today or yesturday; if not stop
    # print (alerts)

    try:
        open("supressMe.json", "r")
    except:
        open("supressMe.json", "w").write("{}").close()

    Supresion = open("supressMe.json", "r")
    SupressMe = json.load(Supresion)

    #print (str(datetime.now().strftime("%Y%m%d")))
    

    UpdatedSupressMe = SupressMe.copy()

    allSupressed = []

    for i in SupressMe:
        for j in SupressMe[i]:
            allSupressed.append(j)

    newSupressMe = []

    MultiAlert = ""
    for i in alerts:
        if i not in allSupressed:
            singleAlert = f"""==============================================================
Serial Number: {i}
ID: {alerts[i]["ID"]}
Customer: {alerts[i]["Customer"]}
Address: {alerts[i]["Address"]}\n"""
            MultiAlert += singleAlert
        newSupressMe.append(i)
    MultiAlert += "=============================================================="

    
    for i in UpdatedSupressMe.copy():
        if int(i) <= int(str(datetime.now() - timedelta(30)).split(" ")[0].replace("-", "")):
            del UpdatedSupressMe[i]
        
    #print (UpdatedSupressMe)

    timeCode = str(datetime.now().strftime("%Y%m%d"))
    #print (timeCode)
    UpdatedSupressMe[timeCode] = newSupressMe

    #print (UpdatedSupressMe)

    with open ("SupressMe.json", "w") as Supresion:
        json.dump(UpdatedSupressMe, Supresion)
        Supresion.close()
    try:
        os.mkdir("output")
    except: pass

    #print (SupressMe)
    timeStamp = str(datetime.now().strftime("%Y%m%d-%H%M%S"))
    fileName = f"output\\Waste Toner List - {timeStamp}.txt"
    with open (fileName, "w") as output:
        output.write(MultiAlert)
        output.close()
    # send email
    print ()

if __name__ == "__main__":
    wasteToner()

# Allow for manual SN additions to suppress me file
  