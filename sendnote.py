#
# Author: Alistair Knox 
#         saknox@ucsd.edu
#
#         May 12, 2018
# 
# Automatically send Endnote emails to the address specified in 
# the command line argument.
#
# Usage:
#
# Open cmd and type the following line, ignoring the '#' and brackets 
# py sendnote.py [license number] [email address]
# where license number is the Endnote number starting with 30881.
# 
# This script assumes an excel format with the Endnote numbers
# in the second column and the keys in the third column.


import sys
import smtplib
from datetime import date
from email.mime.text import MIMEText
import pyexcel as pe

# Error messages
USAGE_MSG = """Usage: py sendnote.py [license number] [email address]
where license number is the Endnote number starting with 30881."""
SAVE_ERR = """Error saving to excel file, another user may be editing it.
To prevent having to manually update the file, no emails were sent.
Close the file and try again."""
SAVE_ERR2 = """Error clearing changes in excel file! No email was sent but the excel file
does not reflect this!"""
UNDO = """Undoing changes to excel file..."""
TERM = "\nProgram terminated. No emails were sent.\n"

# Check number of command line arguments
if len(sys.argv) != 3:
    print(USAGE_MSG + TERM)
    sys.exit()

sender = 'bossmail@ucsd.edu'
password = 'Amazing!'
mailserver = 'smtp.office365.com'
key = 0
recipient = sys.argv[2]
license_num = sys.argv[1]

# Read excel file to get keys
filename='test.xlsx'

sheet = pe.get_sheet(file_name=filename)
print("Searching for license number " + license_num + " in " + filename + "...")
for row in sheet.rows():
    for cell in row:
        if license_num in str(cell):
        # Save the key to send
            key = str(row[2])
        # Write date and recipient to file, unless date is not empty
            if (str(row[3]) != "" or str(row[4]) != ""):
                print("Error: license number "+license_num+" has already been assigned!")
                print("Assigned to "+str(row[4])+" on "+str(row[3])+TERM)
                sys.exit()
            row[3] = date.today().strftime('%m/%d/%Y')
            row[4] = recipient
            break

# Make sure we found a key
if key == 0:
    print("Error: could not find key for "+license_num+" in "+filename+TERM)
    sys.exit()

print("Key found! Updating excel file...")
# Write date and recipient to file
try:
    sheet.save_as(filename)
except:
    print(SAVE_ERR)
    sys.exit()

print("Sending mail...")
# Open a plain text file for the message body.
with open("endnotemsgbody.txt", "r") as f:
    # Create a string to hold the message body
    textbody = f.read()

# Add the key to the message body
keyedmsg = (textbody.replace('[key]',key))

# Create the email message the keyed message body
msg = MIMEText(keyedmsg)
msg['Subject'] = 'Endnote License #' + license_num + ' - UCSD Bookstore'
msg['From'] = sender
msg['To'] = recipient

try:
    # Send the message with the specified SMTP server.
    s = smtplib.SMTP(mailserver,587)
    s.starttls()

    s.login(sender,password)
    s.send_message(msg)
    s.quit()
    print("Success! Activation sent to " + recipient)
except Exception as e:
    print("Email delivery failed\n"+TERM)
    print(e)
    try:
        print(UNDO)
        row[3] = row[4] = ""
        sheet.save_as(filename)
    except:
        print(SAVE_ERR2)


