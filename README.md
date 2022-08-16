# Waste Toner Suppressor
This software is designed to read an emails inbox and picks out the waste toner alerts from Toshiba ECC alerts


# How It Works
1. It initially grabs the congifuration to access the email address that recieves all of the ECC alerts
2. It then downloads the emails that have been sent to the specified inbox folder from the past x amount of days set from the config file
3. it then strips the emails to the few pieces of information needed and adds all alerts to a json file that contains all the alerts from the past 30 days including repeat alerts
4. it then compiles the machines into a txt file with all of the serial numbers that havne't sent an alert in the past 30 days 


# Installation
1. Download the executable
2. Run the executable once
3. With the new config file in the folder structure enter in the email details for the email that all of the alerts get sent to.
4. run the executable again and it will produce a txt file that will contain all of the ECC alerts from the past 4 days.

NOTE: For the first 30 days you may recieve repeat alerts come through as the program builds its database of alerts, the main chunk will come through in the first 4 days and after that it  sould be fine but always double check all of the alerts for best results
