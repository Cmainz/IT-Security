# portMaintenance Script

The sole purpose is to verify that noone has accidentally or intentionally opened a port on your locale computer.

The programme starts of by creating a file with all the ports that are open including the service using the port.
When running it a second time, it will create a hash value of the first file and compare it with new file.
If there is any changes between the original file, an email will be send to person in charge of the pc.
As described below, it is recommended to automate this, like once pr. day or every twelve hours, depending on the needs.

It is not impossible to include more machines for surveliance but as of this moment, the script is missing "try" and "except" statements, and running the script on a whole subnet, might give false emails-  

## Usage

1. Advised to only run the script on a pc that is on your private network, etc. your own home computer.
2. Change __sender_email="xxx.XX@gmail.com"__, __reciever_email="xxx.XX@gmail.com"__ and __server.login("xxx.XXXX@gmail.com",APP password)__ with your own credentials.
3. Might also want to change smtplib.SMTP_SSL(__"SMTP.gmail.com"__ with an email service that you want to use

###### Note
The script has been tested and works if used on linux with cronjob 
but should not be an issue to implement in "Scheduled Tasks" on Windows
