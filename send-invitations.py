#!/usr/bin/env python3
# -*- coding: iso-8859-15 -*-
import csv
from O365 import Account, FileSystemTokenBackend
import readline #for O365 terminal authentication to work with osx 

# Office 365 credentials
client_id = 'changeme'
client_secret = 'changeme'

# member data
members_csv = 'last-year.csv'

#email data
cc_email = "changeme"
body_file = "message-body.html"

# Load member data from CSV
def load_member_data(csv_file):
    members = []
    with open(csv_file, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            members.append((row['first'], row['email']))
    return members

def init_account():
    credentials = (client_id, client_secret)
    account = Account(credentials)
    if not account.is_authenticated:
      if account.authenticate(scopes=['basic', 'message_send']):
        print('Authenticated!')
    return account

def create_body(name):
    body = ""
    with open(body_file, 'r') as file:
        for line in file:
            body += line.replace("{{name}}", name)
    return body

# Send email to each member
def send_email(account, name, email):
        # Create a message object
        message = account.new_message()
        message.to.add(email)
        message.cc.add(cc_email)
        message.subject = "<Subject>"

        message.body = create_body(name)
        # Send the message
        message.send()
        print(f"Sending email to {name} ({email})")

# Main script
if __name__ == '__main__':
    # CSV file path and email details
    members = load_member_data(members_csv)
    account = init_account()
    for (name,email) in members:
      send_email(account, name, email)
