import win32com.client
import os
import pathlib
from datetime import datetime,timedelta
import sys



def retreive_mail_tool(sender_name):
    print(f"Sender : {sender_name}")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)


    messages = inbox.Items
    messages = messages.Restrict(f"[SenderName] = \"{sender_name}\" ")

    print("###############RECUPERATION PJ##########################")

    for message in list(messages):
        try:
            print("========")
            print("Subj: " + message.Subject)
            print("Email Type: ", message.SenderEmailType)
            if message.Class == 43:
                if message.SenderEmailType == "SMTP":
                    print("Name: ", message.SenderName)
                    print("Email Address: ", message.SenderEmailAddress)
                    print("Date: ", message.ReceivedTime)
                elif message.SenderEmailType == "EX":
                    print("Name: ", message.SenderName)
                    print("Sender: ", message.Sender.GetExchangeUser())
                    print("Attachments: ", message.Attachments)
                    print("Email Address: ", message.Sender.GetExchangeUser(
                                                                        ).PrimarySmtpAddress)
                    print("Date: ", message.ReceivedTime)

        except Exception as e:
            print("error when processing emails messages:" + str(e))
            sys.exit(1)
        print("############### FIN DE RECUPERATION PJ ##########################")
    return messages

def save_Attachement(attachment):
    try:        
        attachment.SaveASFile(os.path.join(pathlib.Path().resolve(), f"src/GUI/data/{attachment.FileName}"))    
        print(f"attachment {attachment.FileName} saved")
    except Exception as e:
        print("error when saving the attachment:" + str(e))
        sys.exit(1)