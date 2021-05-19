# portMaintenance Script

The sole purpose is to verify that noone has accidentally or intentionally opened a port on your locale computer.

## Usage

1. Advised to only run the script on a pc that is on your private network, etc. your own home computer.
2. Change __sender_email="xxx.XX@gmail.com"__, __reciever_email="xxx.XX@gmail.com"__ and __server.login("xxx.XXXX@gmail.com",APP password)__ with your own credentials.
3. Might also want to change smtplib.SMTP_SSL(__"SMTP.gmail.com"__ with an email service that you want to use

###### Note
The script has been tested and works if used on linux with cronjob 
but should not be an issue to implement in "Scheduled Tasks" on Windows
