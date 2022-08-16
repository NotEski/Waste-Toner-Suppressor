import EmailHandler
import json, os
from datetime import datetime
from datetime import timedelta
import urllib.request

print ("Loading Files...", end="\r")


defaultConfig = '{"email":"YOUR EMAIL","password":"YOUR EMAILS PASSWORD","server":"outlook.office365.com","port":993,"FromXDaysAgo":4,"KeepFromXDaysAgo":30,"emailFolder":"INBOX"}'


if "config.json" not in os.listdir():
    file = open ("config.json", "w")
    file.write(defaultConfig)
    file.close()
else:
    with open ("config.json", "r") as file:
        config = file.read()
        if "YOUR EMAIL" in config:
            input ("Please fill out the config file and enter your email address and set up SMTP details")
            quit ()


try:
    urllib.request.urlopen('http://google.com')
except:
    print ("ERROR - No Internet Connection")
    input ("Press 'Enter' to exit the program")
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

            
    try:
        open("supressMe.json", "r")                     # Does the Suppresion file exist
    except:
        open("supressMe.json", "w").write("{}").close() # If not create it

    Supresion = open("supressMe.json", "r") # Open the suppresion file
    SupressMe = json.load(Supresion)        # Read the old suppression file


    UpdatedSupressMe = SupressMe.copy()

    allSupressed = []

    for i in SupressMe:
        for j in SupressMe[i]:
            if j not in allSupressed:
                allSupressed.append(j)      # Get a list of all of the old alerts

    newSupressMe = []                       # All of the alerts thatwill come through this run

    MultiAlert = ""
    for i in alerts:
        if i not in allSupressed:           # Compiles a String that contains information for the alerts file 
            singleAlert = f"""{("="*50)}
Serial Number: {i}
ID: {alerts[i]["ID"]}
Customer: {alerts[i]["Customer"]}
Address: {alerts[i]["Address"]}\n"""
            MultiAlert += singleAlert
        newSupressMe.append(i)
    MultiAlert += ("="*50)

    
    for i in UpdatedSupressMe.copy():
        if int(i) <= int(str(datetime.now() - timedelta(30)).split(" ")[0].replace("-", "")):  # Checks if the entry in the suppresion file is older then 30 days
            del UpdatedSupressMe[i]                                                            # Deletes it if its too old
        

    timeCode = str(datetime.now().strftime("%Y%m%d"))   # Gets the datecode for adding the old alerts into the suppresion folder
    UpdatedSupressMe[timeCode] = newSupressMe           # Adds List of new alerts against the timecode into the new file 

    with open ("SupressMe.json", "w") as Supresion:     # Writes the new suprpesion file
        json.dump(UpdatedSupressMe, Supresion)
        Supresion.close()
    try:
        os.mkdir("output")      # Makes a new folder output for all of the alert txt files
    except: pass

    timeStamp = str(datetime.now().strftime("%Y%m%d-%H%M%S"))   # Creates the time stamp that contains all of the 
    fileName = f"output\\Waste Toner List - {timeStamp}.txt"
    with open (fileName, "w") as output:
        output.write(MultiAlert)
        output.close()
    # send email
    print ()

if __name__ == "__main__":
    wasteToner()

# Allow for manual SN additions to suppress me file
  
