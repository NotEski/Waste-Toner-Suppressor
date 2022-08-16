# Waste Toner Suppressor
This software is designed to read an emails inbox and picks out the waste toner alerts from Toshiba ECC alerts


# How It Works
1. It initially grabs the congifuration to access the email address that recieves all of the ECC alerts
2. It then downloads the emails that have been sent to the specified inbox folder from the past x amount of days set from the config file
3. it then strips the emails to the few pieces of information needed and adds all alerts to a json file that contains all the alerts from the past 30 days including repeat alerts
4. it then compiles the machines into a txt file with all of the serial numbers that havne't sent an alert in the past 30 days 
