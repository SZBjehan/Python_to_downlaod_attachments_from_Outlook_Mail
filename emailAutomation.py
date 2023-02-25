import win32com.client
import re               
from datetime import datetime

outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")

# replace email@email.com with your outlook email
inbox = outlook.Folders('email@email.com').Folders('Inbox')

# from which date you want to start downloading (yyyy, mm, dd)
# start_date = datetime(2023, 2, 4)
current_date = datetime.date.today()

messages = inbox.items

for msg in messages:
    # replace Report title with common text in the subject line
    if "Report title" in msg.Subject:
        # Extract the date from the string
        date_str = msg.Subject.split()[-1]
        
        if current_date in msg.Subject:
            # this is to create a folder where the attachments will be saved
            if not os.path.exists(‘Folder_Name’):
                os.makedirs(‘Folder_Name’)

            for attachment in msg.Attachments:
                #downloading and storing the attachments
                attachment.SaveAsFile(os.getcwd() + ‘\\Folder_Name\\’ + attachment.FileName)