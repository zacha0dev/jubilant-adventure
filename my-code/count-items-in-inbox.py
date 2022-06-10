#
# Outlook Automation 
# Count number of items in your Inbox  
#
# Author: Zachary Allen
# Date: June 9, 2022
#

#
# Dependencies: pywin32 (https://pypi.org/project/pywin32/) 
#
# Run the below commands: 
# pip install pywin32
# 

from email import message 
import win32com.client, os, win32ui

# variables
email_one = 'email_one'  # Replace with you email from outlook
email_two = 'email_two'  # If you have a second one you can add it here 

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # Create a reference to Outlook 
account_one = outlook.Folders[email_one]  # Make a reference to the folders of account 1
account_two = outlook.Folders[email_two]  # Make a reference to the folders of account 2

inbox_one = account_one.Folders['Inbox']  # Make a reference to the inbox of account 1
inbox_two = account_two.Folders['Inbox']  # Make a reference to the inbox of account 2

print(inbox_one.Items.Count)  # Returns the count of items of inbox of account 1
print(inbox_two.Items.Count)  # Returns the count of items of inbox of account 2


