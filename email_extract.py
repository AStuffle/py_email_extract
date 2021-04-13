import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Target the folder you want to pull from. Use the get_outlook_folders script to get a complete list and re-target as needed.
inbox = outlook.GetDefaultFolder(6).Folders[15]

# 
messages = inbox.Items
message = messages.GetFirst()

# Attachment save location. Needs to exist already.
filepath = 'C:\\Users\\andrews.BROWN\\OneDrive - Brown Distributing Company, LTD\\NAA\\IRI Reports\\Exports'

for m in messages:
    print (message)
    attachments = message.Attachments
    attachment = attachments.Item(1)
    print (attachment)
    
    # Comment this out until you're sure you know what is going to get pulled!
    # This will currently pull the first attachment from every email in the target folder!
    #attachment.SaveASFile(os.path.join(filepath,attachment.FileName))
    
    message = messages.GetNext()
