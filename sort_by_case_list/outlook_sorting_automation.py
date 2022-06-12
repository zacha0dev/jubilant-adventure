#
# Outlook Sorting Automation  
#
# Author: Zachary Allen
# Release Date: June 12, 2022
# Version 1 
#

#
# Dependencies: pywin32 (https://pypi.org/project/pywin32/) 
#
# Run the below commands: 
# pip install pywin32
#
# You can install pywin32 into your project folder by using: 'py -m pip install pywin32'
# 

import win32com.client

# creates references for outlook and email accounts  
def set_admin():
    # sets email variables 
    email_one = 'email_one'  # specify the email account to use
    email_two = 'email_two'  # specify a second email account if you have one 

    # sets an outlook reference 
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # sets account references    
    account_list = []
    account_list.append(outlook.Folders[email_one])  # account_list[0]
    account_list.append(outlook.Folders[email_two])  # account_list[1]
    return  account_list

# creates folder references for each account in account_list 
def set_folders(account_list):
    folder_list = []
    folder_list.append(account_list[0].Folders['Inbox'])  # folder_list[0]
    folder_list.append(account_list[1].Folders['Inbox'])  # folder_list[1]
    folder_list.append(account_list[1].Folders['Inbox'].Folders.Item('Cases'))  # folder_list[2]
    return folder_list

# gets list of subfolders from the parent folder passed into the function 
def get_subfolders(parent_folder):
    subfolders = []
    for i in range(1, (parent_folder.Folders.Count + 1), 1):
        subfolders.append(str(parent_folder.Folders.Item(i)))
    return subfolders  # returns list of folders under the passed in parent folder 

# gets list of data from file 
def get_file():
    file_path = "/"  # path to the case_list.txt file - remember to use "/" for each path rather than windows default of "\" 

    try: 
        my_file = open(file_path, "r")
        return my_file.read().split("\n")
    except Exception:
        print("Error - File does not exist: '" + file_path + "'\r\n" + "Correct file path in Func: get_case_list()\r\n")  # throws an error if the file cannot be found
        quit()  # quits the script if the file cannot be found 

# checks that folders are created based on case_list.txt 
def verify_cases_are_created(cases, folder):
    for i in range(0, len(cases), 1):
        try: 
            folder.Folders.Item(str(cases[i]))
        except Exception:
            folder.Folders.Add(cases[i])
            continue

# deletes any other subfolders that are not in the case_list.txt 
def delete_folders_not_in_case_list(list_to_delete, folder):
    for i in range(0, len(list_to_delete), 1):
        try: 
            folder.Folders.Item(str(list_to_delete[i]))
            folder.Folders.Item(list_to_delete[i]).Delete()
        except Exception:
            continue
    return False

# sorts the messages from the inbox into the respected destination subfoler if the subject name includes the case number from case_list.txt
def sort_messages(inbox, destination_folder, cases):
    emails_to_move = []

    for messages in inbox.Items: 
        for i in range(0, len(cases), 1):
            if str(messages).find(cases[i]) != -1:
                emails_to_move.append(messages)
                #messages.Move(destination_folder.Folders.Item(cases[i]))

    for messages in emails_to_move:
        for i in range(0, len(cases), 1):
            if str(messages).find(cases[i]) != -1:
                messages.Move(destination_folder.Folders.Item(cases[i]))

# Start of Program  

accounts = set_admin()
folder = set_folders(accounts)
subfolders = get_subfolders(folder[2])
cases = get_file()

verify_cases_are_created(cases, folder[2])
delete_folders_not_in_case_list(list(set(subfolders) - set(cases)), folder[2])
sort_messages(folder[0], folder[2], cases)
sort_messages(folder[1], folder[2], cases)

