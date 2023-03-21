import win32com.client
import os
import pathlib
from datetime import datetime,timedelta
import sys



def retreive_mail_tool():
    name_attachement = ""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    received_dt = datetime.now() - timedelta(hours=1)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = inbox.Items
    messages = messages.Restrict(f"[SenderName] = 'TLECORNE@bouyguestelecom.fr'")
    outputDir = r"C:\Users\tbonnard\Documents\GIT\Repositories\automatize_outlook_maildata\data"

    print("###############RECUPERATION PJ##########################")

    
    for message in list(messages):
    
    # list(messages)[len(list(messages))-26]
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
            try:        
                s = message.sender
                print(type(message.Attachments))
                for attachment in message.Attachments:
                    name_attachement = attachment.FileName
                    attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                    print(f"attachment {attachment.FileName} from {s} saved")
            except Exception as e:
                print("error when saving the attachment:" + str(e))
                sys.exit(1)
        except Exception as e:
            print("error when processing emails messages:" + str(e))
            sys.exit(1)
        print("############### FIN DE RECUPERATION PJ ##########################")
    return name_attachement


retreive_mail_tool()