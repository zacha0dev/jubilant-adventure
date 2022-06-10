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

# creates outlook and email account references 
def set_admin():
    # Set Email Variables
    email_one = 'email_one'  # Replace with you email from outlook
    email_two = 'email_two'  # If you have a second one you can add it here

    # Create Outlook reference
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Set accounts   
    accounts = []
    account_one = outlook.Folders[email_one]  # Make a reference to the folders of account 1
    account_two = outlook.Folders[email_two]  # Make a reference to the folders of account 2
    accounts.append(account_one) 
    accounts.append(account_two)
    return accounts  # Returns a list of the accounts 

# set folder paths
def set_folders(accounts):
    folders = []
    inbox_one = accounts[0].Folders['Inbox']  # Make a reference to the inbox of account 1
    inbox_two = accounts[1].Folders['Inbox']  # Make a reference to the inbox of account 2
    cases_folder = inbox_two.Folders.Item('Cases')   # Make a reference to a subfolder 'Cases' of inbox_two, this can be any subfolder
    folders.append(inbox_one)
    folders.append(inbox_two)
    folders.append(cases_folder)
    return folders  # Returns a list of each of the folders

# gets list of subfolders from the parent folder passed into the function 
def get_subfolders(folder):
    subfolders = []
    subfolder_count = folder.Folders.Count + 1    # Gets a coutn of the total amount of subfolder under the parent folder passed in with an index offset
    for i in range(1, subfolder_count, 1):        # Iterates through each of the subfolders
        subfolder = folder.Folders.Item(i)        # Saves the folder name
        subfolders.append(subfolder)              # Appends the folder name to the list
    return subfolders                             # Returns the list of subfolders

accounts = set_admin()
folders = set_folders(accounts)
subfolders = get_subfolders(folders[2])

# prints out list of subfolders based on you selection
for i in range(len(subfolders)):
     print(subfolders[i])
