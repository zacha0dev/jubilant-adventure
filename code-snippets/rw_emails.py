import win32com.client
import unicodecsv as csv
import os

output_file = open('./outlook_farming_001.csv','wb')
output_writer = csv.writer(output_file, delimiter=';', encoding='latin2')

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders.Item("Security Availabilities")
messages = inbox.Items

for i, message in enumerate(messages):
    try:

        sender = message.SenderName
        sender_address = message.SenderEmailAddress
        sent_to = message.To
        date = message.LastModificationTime
        subject = message.subject
        body = message.body

        attachments = message.Attachments
        attachment = attachments.Item(1)
        for attachment in message.Attachments:
            attachment.SaveAsFile(os.path.join(output_file, str(attachment)))

        output_writer.writerow([
            sender,
            sender_address,
            subject,
            body,
            attachment])

    except Exception as e:
        ()

output_file.close()

# Source: https://stackoverflow.com/questions/55654922/use-python-to-connect-to-outlook-and-read-emails-and-attachments-then-write-them
