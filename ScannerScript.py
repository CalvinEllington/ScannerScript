# -*- coding: utf-8 -*-
"""
Created on Fri Jan 18 14:52:34 2019

@author: cellington
"""

import win32com.client
import os
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
message = messages.GetFirst()
subject = message.Subject
#
get_path = '*Place your path here*'

for m in messages:
    if m.Subject == "*REDACTED FOR PORTFOLIO SHOWCASE*":
        print(message)
        message.Attachments
        for attachment in message.Attachments:
            attachment.SaveASFile(os.path.join(get_path,attachment.FileName))
            print(attachment)
        message = messages.GetNext()

    else:
        message = messages.GetNext()
