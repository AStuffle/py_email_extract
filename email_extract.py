import win32com.client
import os
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders[15]
messages = inbox.Items
message = messages.GetFirst()

filepath = 'C:\\Users\\andrews.BROWN\\OneDrive - Brown Distributing Company, LTD\\NAA\\IRI Reports\\Exports'

for m in messages:
    print (message)
    attachments = message.Attachments
    attachment = attachments.Item(1)
    print (attachment)
    #attachment.SaveASFile(os.path.join(filepath,attachment.FileName))
    message = messages.GetNext()
