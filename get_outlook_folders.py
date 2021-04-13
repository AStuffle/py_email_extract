# This script can be used to get folder object references from outlook.
# Run it from cmd or shell.

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

print('\n\n')
input('Press enter to list Outlook folders and their indexes...')
print('\n\n')

getfolders()
print('\n\n')

parentfolder = input('Select a folder number from above to see subfolders: ')
print('\n\n')

getsubfolders(parentfolder)
print('\n\n')

input('Press enter to exit.')
