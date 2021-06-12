#!/usr/bin/python3

import nmap
import sys
import os
import hashlib
import smtplib, ssl
import argparse

#################################

parser = argparse.ArgumentParser(description='Please specify a target IP')


parser.add_argument('ip_input', metavar='IP', type=str,
                    help='input an ip to be reviewed')

args = parser.parse_args()

######################################

def show_hash(filename):
  hashed=hashlib.sha1()
  with open(filename, 'rb') as file:
    chunk=0
    while chunk !=b'':
      chunk=file.read(1024)
      hashed.update(chunk)

  return hashed.hexdigest()

def printToFile(): #Creates a log file
  record = open(host_file,"w+")

  record.write("""
  Host: {} was scanned.
  The following ports are open:\n
  """.format(host))

  for proto in nm[host].all_protocols():
    record.write("-----------------\n\n")
    record.write("Protocol : {}\n".format(proto))
    for ports in nm[host][proto].keys():
      serviceName=nm[host]['tcp'][ports]['name']

      record.write( "Port : {} service: {}".format(ports,serviceName))
  record.close()


#####################################3

nm=nmap.PortScanner()
host =str(sys.argv[1])
host_file=host+'.txt'
nm.scan(host,arguments='-sV sS')



if os.path.exists(host_file):
  oldRecordHash =show_hash(host_file)

  printToFile() #Overwrites old file, but keeps the SHA1 value

  newRecordHash =show_hash(host_file)

  if oldRecordHash == newRecordHash:
    print("There seems to be nothing to report")
  else:

    context = ssl.create_default_context()
    sender_email="xxx.XX@gmail.com"
    reciever_email="xxx.XX@gmail.com"
    content=open(host_file, 'r')
    message ="Subject: WARNING - Review of log is needed\n\n {}".format(content.read())
    with smtplib.SMTP_SSL("SMTP.gmail.com",context=context) as server:
      server.login("xxx.XXXX@gmail.com","xxXXxx")
      server.sendmail(sender_email,reciever_email,message)
      # Send an email if any changes

else:
  printToFile()
  print("""
  Creating a new log. If this is not first time running the script,
  please review your ports and file
  """)
