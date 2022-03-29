from fileinput import filename
import win32com.client
#other libraries to be used in this script
import openpyxl
from openpyxl import Workbook
import jsonlines

emails = {}


outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
errors = []

for account in mapi.Accounts:
	print(account.DeliveryStore.DisplayName)
print("\n")
inbox = mapi.GetDefaultFolder(6)
folderFollet = mapi.GetDefaultFolder(6).Folders["This Week Mr. Follet"]
messages = folderFollet.Items
for message in list(messages):
    fileName = message.subject
    if "/" in message.subject:
        fileName = ""
        for bit in message.subject.split("/"):
            fileName = fileName + bit + "."
        fileName = fileName[0:len(fileName)-1:]
        #print("Modified: " + fileName)
    else:
        x=False
        #print(fileName)
    contentFile = open("Mr. Follet Emails\\" + fileName + ".txt", "w", encoding='utf-8')
    body = ""
    for pieces in message.Body.split("\n"):
        if pieces.strip() != "" and pieces.strip() != "\n" and pieces.strip() != " \n":
            #print("|" + pieces.strip() + "|")
            body = body + pieces.strip() + "\n"
    emails[fileName] = body
    contentFile.write(body)
    contentFile.close()
with jsonlines.open('output.jsonl', mode='w') as writer:
    for subjects in emails.keys():
        writer.write('{"prompt": "' + subjects + '##SEPERATION##", "completion": "' + emails[subjects] + '"}')

print("Completed!")