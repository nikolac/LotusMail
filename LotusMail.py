from win32com.client import Dispatch
import datetime

"""
----- Dependencies ----

pyWin (win32com)
sourceforge.net/projects/pywin32/

----- Example Usage ----

from LotusMail import Mail

lMail = Mail("mypassword01", "", "mail3\\myusername.nsf")

lMail.send("toemail@email.com", "emailsuject line", "Dear sender, ... Yours truely, usernname", ['c:\username\mydoc',])

"""

class Mail:
    def __init__(self, password, server, mailFile):
        self.password = password
        self.server = server
        self.mailFile = mailFile
        self.session = Dispatch('Lotus.NotesSession')
        self.session.Initialize(self.password)
        self.db = self.session.getDatabase(self.server, self.mailFile)
        


    def send(self, to, subject, body, attachFiles = []):
        mail = self.db.CREATEDOCUMENT()
        mail.ReplaceItemValue("Form", "Memo")
        mail.ReplaceItemValue("Subject", subject)
        mail.ReplaceItemValue("SendTo", to)
        mailBody = mail.CREATERICHTEXTITEM("Body")
        mailBody.APPENDTEXT(body)
        mailBody.ADDNEWLINE(2)
        
        for attachFile in attachFiles:
            mailBody.EMBEDOBJECT(1454, "", attachFile, "Attachment")
            mailBody.ADDNEWLINE(2)
            
        mail.ReplaceItemValue("PostedDate", datetime.datetime.now())
        mail.SAVEMESSAGEONSEND = True
        mail.SEND(False)
        
    

         
