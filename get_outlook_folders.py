# This script can be used to get folder object references from outlook.
# Best to run it from comand line or shell.

import win32com.client
import os
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def getfolders():
    for i in range(100):
        try:
            box = outlook.GetDefaultFolder(i)
            name = box.Name
            print(i, name)
        except:
            pass

def getsubfolders(parentfolder):
    for i in range(100):
        try:
            box = outlook.GetDefaultFolder(parentfolder).Folders[i]
            name = box.Name
            print(i, name)
        except:
            pass

print('Ready!')
print('Use getfolders() to get a list of top-level folders and their index numbers.')
print('Use getsubfolders(6) to get subfolders of folder 6 (Inbox)').
