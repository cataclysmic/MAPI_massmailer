# -*- coding: utf-8 -*-

#!/usr/bin/env python3

# import tkinter as tk
from tkinter import Tk, Label, Button, StringVar, Entry, filedialog, W, E
from tkinter import Radiobutton, Checkbutton, IntVar, Frame
from tkinter import messagebox as mbox
from time import sleep
from time import strftime, ctime
from hashlib import md5

from os import listdir
#import pandas as pd
from pandas import read_excel
from win32com.client import Dispatch


class CoreGui(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)

        self.attachFields = [] # stores UI attachment fields
        self.attachments = {} # stores attachment references
        self.recConf = IntVar()

        # labels
        self.Lang_From = StringVar()
        self.Lang_Subject = StringVar()
        self.Lang_Recipient = StringVar()
        self.Lang_uId = StringVar()
        self.Lang_Mail = StringVar()
        self.Lang_Html = StringVar()
        self.Lang_Text = StringVar()
        self.Lang_Attach = StringVar()
        self.Lang_Format = StringVar()
        self.Lang_Receipt = StringVar()
        self.Lang_Interval = StringVar()
        self.Lang_Address = StringVar()
        self.Lang_Content = StringVar()
        self.Lang_Parse = StringVar()
        self.Lang_Send = StringVar()
        self.Lang_Help = StringVar()
        self.Lang_Log = StringVar()
        self.Lang_File = StringVar()
        self.Lang_Folder = StringVar()

        # set text information - English
        self.SetEng()

        self.master = master
        master.title("Outlook / Exchange Massmailer")

        self.senderMailLab = Label(master, textvariable=self.Lang_From,width=20)
        self.senderMail = Entry(master, width=40)

        self.subjectLab = Label(master, textvariable=self.Lang_Subject)
        self.subject = Entry(master, width=40)

        self.recipientFileSel = Button(master, textvariable=self.Lang_Recipient,
                                       command=self.LoadRecipient, bg="#6495ed")

        self.recipientFileStr = StringVar()
        self.recipientFile = Entry(master, width=40, state='readonly',
                                   textvariable=self.recipientFileStr)

        self.uniqueIdLab = Label(master, textvariable=self.Lang_uId)
        self.uniqueId = Entry(master, width=40)

        self.mailIdLab = Label(master, textvariable=self.Lang_Mail)
        self.mailId = Entry(master, width=40)

        self.mailBodyHtmlSel = Button(master, textvariable=self.Lang_Html,
                                      command=self.LoadBodyHtml, bg="#6495ed")
        self.mailBodyHtmlStr = StringVar()
        self.mailBodyHtmlStr.set("")
        self.mailBodyHtml = Entry(master, width=40, state='readonly',
                                  textvariable=self.mailBodyHtmlStr)

        self.mailBodySel = Button(master, textvariable=self.Lang_Text,
                                  command=self.LoadBodyText, bg="#6495ed")
        self.mailBodyStr = StringVar()
        self.mailBodyStr.set("")
        self.mailBody = Entry(master, width=40, state='readonly',
                              textvariable=self.mailBodyStr)

        self.AttachmentsLab = Label(master, textvariable=self.Lang_Attach)
        self.addAttachment = Button(master, text="+",
                                    command=self.AddAttachmentField, bg="#6495ed")

        self.mailFormatLab = Label(master, textvariable=self.Lang_Format)
        self.mailForm = IntVar()
        self.mailForm.set(3)
        self.mailFormat1 = Radiobutton(master, text="Text",
                                       variable=self.mailForm, value=1)
        self.mailFormat2 = Radiobutton(master, text="Html",
                                       variable=self.mailForm, value=2)
        self.mailFormat3 = Radiobutton(master, text="Text & Html",
                                       variable=self.mailForm, value=3)

        self.receiptConfirm = Checkbutton(master, textvariable=self.Lang_Receipt,
                                          variable=self.recConf)

        self.space = Label(master,text=" ")
        self.space1 = Label(master,text=" ")
        self.space2 = Label(master,text=" ")
        self.address = Label(master, textvariable=self.Lang_Address)
        self.mailtext = Label(master, textvariable=self.Lang_Content)

        self.v = IntVar()
        self.pauseLab = Label(master, textvariable=self.Lang_Interval)
        self.pause = Entry(master, textvariable=self.v, width=4)
        self.v.set(5)

        self.parse = Button(master, textvariable=self.Lang_Parse, command=self.ParseMail, width=50, bg="#6495ed")
        self.send = Button(master, textvariable=self.Lang_Send, command=self.SendMail, width=50, bg="#6495ed")
        self.help = Button(master, textvariable=self.Lang_Help, command=self.onInfo, bg="#6495ed")

        self.LangEn = Button(master, text="EN", command=self.SetEng, bg="#6495ed")
        self.LangDe = Button(master, text="DE", command=self.SetDe, bg="#6495ed")

        self.logLocationStr = StringVar()
        self.logLoc = Button(master, textvariable=self.Lang_Log, command=self.LogFolder, bg="#6495ed")
        self.logLocation = Entry(master, width=40, state='readonly',
                                 textvariable=self.logLocationStr)

        # layout the program
        self.LangEn.grid(row=0, column=0, sticky=E)
        self.LangDe.grid(row=0, column=1, sticky=W)
        self.help.grid(row=0, column=1, sticky=E)
        self.senderMailLab.grid(row=1, column=0, sticky=E)
        self.senderMail.grid(row=1, column=1,sticky="EW")
        self.subjectLab.grid(row=3, column=0, sticky=E)
        self.subject.grid(row=3, column=1, sticky="EW")
        self.space.grid(row=4, columnspan=2)
        self.address.grid(row=5, columnspan=2, sticky="EW")
        self.recipientFileSel.grid(row=6, column=0, sticky=E)
        self.recipientFile.grid(row=6, column=1, sticky="EW")
        self.uniqueIdLab.grid(row=7, column=0, sticky=E)
        self.uniqueId.grid(row=7, column=1, sticky="EW")
        self.mailIdLab.grid(row=8, column=0, sticky=E)
        self.mailId.grid(row=8, column=1, sticky="EW")
        self.space1.grid(row=9, columnspan=2)
        self.mailtext.grid(row=10, columnspan=2)
        self.mailFormatLab.grid(row=11, column=0)
        self.mailFormat1.grid(row=11, column=1, sticky=W)
        self.mailFormat2.grid(row=11, column=1)
        self.mailFormat3.grid(row=11, column=1, sticky=E)
        self.mailBodyHtmlSel.grid(row=12, column=0, sticky=E)
        self.mailBodyHtml.grid(row=12, column=1, sticky="EW")
        self.mailBodySel.grid(row=13, column=0, sticky=E)
        self.mailBody.grid(row=13, column=1, sticky="EW")
        self.receiptConfirm.grid(row=15, column=1, sticky=W)
        self.AttachmentsLab.grid(row=16, columnspan=2)
        self.addAttachment.grid(row=16, column=1, sticky=E)

        self.grid(columnspan=2, sticky="NEWS")

        self.parse.grid(columnspan=2, sticky="EW")
        self.logLoc.grid(row=20, column=0, sticky=E)
        self.logLocation.grid(row=20, column=1, sticky="EW")
        self.pauseLab.grid(row=21, column=0, sticky=E)
        self.pause.grid(row=21, column=1, sticky=W)
        self.send.grid(columnspan=2, sticky="EW")

    def LogFolder(self):
        '''select log file folder location'''

        foldername = filedialog.askdirectory()
        self.logLocationStr.set(foldername)

    def LoadRecipient(self):
        '''Load spreadsheet file with recipients.'''

        filename = filedialog.askopenfilename(
            filetypes=(("XLSX", "*.xlsx"),
                       ("All files", "*.*")))

        self.recipientFileStr.set(filename)

        self.recipientDf = read_excel(filename)
        print(self.recipientDf)

    def LoadBodyText(self):
        '''Load html file contain mail body'''

        filename = filedialog.askopenfilename(
            filetypes=(("Text", "*.txt"),
                       ("Org", "*.org"),
                       ("MarkDown", "*.md")))

        self.mailBodyStr.set(filename)

        file = open(filename, 'r')
        self.mailBodyRaw = file.read()
        self.mailBodyRaw = self.mailBodyRaw.replace("[", "{").replace("]", "}")

    def LoadBodyHtml(self):
        '''Load html file contain mail body'''

        filename = filedialog.askopenfilename(
            filetypes=(("HTML", "*.html"),
                       ("HTM", "*.htm")))

        self.mailBodyHtmlStr.set(filename)

        file = open(filename, 'r')
        self.mailBodyHtmlRaw = file.read()
        self.mailBodyHtmlRaw = self.mailBodyHtmlRaw.replace("[", "{").replace("]", "}")

    def AddAttachmentField(self):
        '''add attachment via + '''

        n = len(self.attachFields)
        self.attachFields.append({})

        self.attachFields[n]['label'] = Label(self, text="("+str(n+1)+")")
        self.attachFields[n]['space'] = Label(self, text="   ")

        self.attachFields[n]['folderBut'] = Button(self, textvariable=self.Lang_Folder,
                                                   command=lambda: self.LoadAttachFolder(n),bg="#6495ed")
        self.attachFields[n]['fileBut'] = Button(self, textvariable=self.Lang_File,
                                                 command=lambda: self.LoadAttachFile(n),bg="#57cefa")
        self.attachFields[n]['removeBut'] = Button(self, text="-",
                                                   command=lambda: self.RemoveFile(n),bg="#cd5c5c")
        self.attachFields[n]['stringVar'] = StringVar()
        self.attachFields[n]['field'] = Entry(self, width=43, state="readonly",
                                              textvariable=self.attachFields[n]['stringVar'])


        self.attachFields[n]['fileBut'].grid(row=n, column=0, sticky=W)
        self.attachFields[n]['folderBut'].grid(row=n, column=1, sticky=W)
        self.attachFields[n]['space'].grid(row=n, column=2)
        self.attachFields[n]['field'].grid(row=n, column=3, sticky="EW")
        self.attachFields[n]['label'].grid(row=n, column=4, sticky="E")
        self.attachFields[n]['removeBut'].grid(row=n, column=5, sticky="E")

    def LoadAttachFile(self, idNr):
        '''load attachment file path'''

        filename = filedialog.askopenfilename(
            filetypes=(("All files", "*.*"),
                       ("PDF", "*.pdf")))

        self.attachFields[idNr]['stringVar'].set(filename)
        self.attachments[idNr] = filename

        print(self.attachments)

    def LoadAttachFolder(self, idNr):
        '''load folder path and list'''

        foldername = filedialog.askdirectory()
        self.attachFields[idNr]['stringVar'].set(foldername)

        fileList = listdir(foldername)
        self.attachments[idNr] = [foldername, fileList]

        print(self.attachments)

    def RemoveFile(self, idNr):
        '''removing attached files'''

        self.attachFields[idNr]['stringVar'].set('')

        self.attachments.pop(idNr)

        print(self.attachments)

    def onInfo(self):
        '''helpful message '''
        mbox.showinfo("Nutzung", "%s" % self.hint)

    def ParseMail(self):
        '''collate all necessary data for single email and send'''

        try:
            senderMail = self.senderMail.get()
            subject = self.subject.get()
            recipient = self.recipientDf
            uniqueId = self.uniqueId.get()
            mailId = self.mailId.get()
            attach = self.attachments

        except AttributeError:
            mbox.showwarning("Error", "%s" % self.Lang_FormError)
            return()


        if int(self.mailForm.get()) == 1 and self.mailBodyStr.get().strip() == "":
            mbox.showwarning("Error", "%s" %  self.Lang_TError)
            return()
        elif int(self.mailForm.get()) == 2 and self.mailBodyHtmlStr.get().strip() == "":
            mbox.showwarning("Error", "%s" % self.Lang_HError)
            return()
        elif int(self.mailForm.get()) == 3 and (self.mailBodyHtmlStr.get().strip() == "" or self.mailBodyStr.get().strip() == ""):
            mbox.showwarning("Error", "%s" % self.Lang_THError)
            return()

        mailForm = int(self.mailForm.get())

        OutBox = self.GetOutbox()

        for i in range(0, len(recipient.index)):
            replacement = recipient.iloc[[i]].to_dict('records')
            print(replacement)
            if mailForm == 1:
                mailBody = self.mailBodyRaw
                bodyFormatTxt = mailBody.format(**replacement[0])
            elif mailForm == 2:
                mailBodyHtml = self.mailBodyHtmlRaw
                #mailBodyHtml = mailBodyHtml.split("<body ")[1].split("\n", 1)[1].split("</body>")[0]
                bodyFormatHtml = mailBodyHtml.format(**replacement[0])
            else:
                mailBody = self.mailBodyRaw
                mailBodyHtml = self.mailBodyHtmlRaw
                #mailBodyHtml = mailBodyHtml.split("<body ")[1].split("\n", 1)[1].split("</body>")[0]
                bodyFormatTxt = mailBody.format(**replacement[0])
                bodyFormatHtml = mailBodyHtml.format(**replacement[0])

            # create message
            #Msg = OutBox.Items.Add(0)
            Msg = self.GetOutbox().Items.Add(0)
            Msg.Subject = subject
            Msg.To = replacement[0][mailId].replace("\n","").strip()
            Msg.SentOnBehalfOfName = senderMail

            if mailForm == 1:
                Msg.Body = bodyFormatTxt.encode("cp1252")
            elif mailForm == 2:
                Msg.HTMLBody = bodyFormatHtml.encode("cp1252")
            else:
                Msg.Body = bodyFormatTxt.encode("cp1252")
                Msg.HTMLBody = bodyFormatHtml.encode("cp1252")

            # Read receipt
            if int(self.recConf.get()) == 1:
                Msg.ReadReceiptRequested = True

            # add attachments
            for j in range(0, len(attach)):
                # single file
                if type(attach[j]) is str:
                    Msg.Attachments.Add('"'+attach[j]+'"')
                # multiple files from folder
                elif type(attach[j]) is list:
                    matches = [s for s in attach[j][1] if str(replacement[0][uniqueId]) in s]
                    if len(matches) > 0:
                        for el in matches:
                            addAttach = '"'+attach[j][0]+'/'+el+'"'
                            Msg.Attachments.Add(addAttach)

            # send message
            # Msg.display()
            Msg.Save()

    def SendMail(self):
        '''iterate through folder and send mails'''

        logPath = self.logLocation.get()+"/"+"massmail_log_"+strftime("%Y%m%d-%H%M%S")

        file = open(logPath+".log", "a")
        file.write(self.subject.get()+"\n\n")
        OutBox = self.GetOutbox()
        messages = OutBox.Items
        print(messages)
        for i in range(0, len(messages)):
            msg = messages.GetLast()
            toMail = str(msg.To)
            file.write(str(ctime())+" - "+toMail+"\n")
            if (self.mailForm.get()) == 2:
                contMail = str(msg.HTMLBody)
            else:
                contMail = str(msg.Body)
            file.write(contMail+"\n")
            file.write("------------------------------------------------------\n")
            msg.Send()
            sleep(float(self.pause.get()))

        file.close()

        # md5 fingerprint of log file
        file = open(logPath+".md5", "w")
        logFile = open(logPath+".log", "rb")
        lF = logFile.read()
        file.write(md5(lF).hexdigest())
        file.close()

    def GetOutbox(self):
        '''find the storage folder'''

        fromMail = self.senderMail.get()

        appliOut = Dispatch("Outlook.Application").GetNamespace("MAPI")
        for i in range(1, 20):
            accounts = appliOut.Folders(i)
            if str(accounts) == fromMail:
                for j in range(1, 20):
                    boxes = appliOut.Folders(i).Folders(j)
                    if str(boxes) == "Entwürfe" or str(boxes) == "Drafts":
                        break
                break
        return(boxes)

# --- English
    def SetEng(self):
        '''set english'''

        # labels
        self.Lang_From.set("From:")
        self.Lang_Subject.set("Subject:")
        self.Lang_Recipient.set("Recipient file:")
        self.Lang_uId.set("ID column:")
        self.Lang_Mail.set("Email column:")
        self.Lang_Html.set("HTML file:")
        self.Lang_Text.set("Text file:")
        self.Lang_Attach.set("Add Attachments:")
        self.Lang_Format.set("Format:")
        self.Lang_Receipt.set("Require read receipt")
        self.Lang_Interval.set("Send interval in sec.:")
        self.Lang_Address.set("Bulk mail data")
        self.Lang_Content.set("Email content")
        self.Lang_Parse.set("" "(1) Create email drafts")
        self.Lang_Send.set("(2) Send emails" )
        self.Lang_Help.set("Hints")
        self.Lang_Log.set("Log folder:")
        self.Lang_File.set("File:")
        self.Lang_Folder.set("Folder:")
        self.Lang_TError = "The chosen format requires the selection of a text file."
        self.Lang_HError = "The chosen format requires the selection of an html file."
        self.Lang_THError = "The chosen format requires the selection of a text and an html file."
        self.Lang_FormError = "Please check the form.\nNot all required fields are filled."
        self.hint = """
          'From:' - Sender email
       'Subject:' - Email subject

            Bulk mail data
'Recipient file' - XLSX-File containing bulk mail data.
                   Data has to be stored in the first spreadsheet.
     'ID column' - Column name (1. row) identifying each recipients IDs.
                   This ID is used to select attachments from selections.
  'Email column' - Column name identifying recipient email.
'Require read receipt'   - Activates the read receipt request.

            Attachments
Using the "+" symbol you can add an arbitrary number of attachments.
         'File' - Single file that gets attached to each email.
       'Folder' - File collection from which a file gets selected for attaching based on ID.

            Email content
Emails can be send as text, html or both. You need to provide the appropriate files.
You can creates those using, e.g., Word. The tool does a character replacement in the mail
body similar to a bulk letter program. The replacement data is contained in
additional columns in the recipient file. Replacement texts is marked by {Column name}.

            'Log-folder' - Folder to save the log file to.
'Send interval in sec.:' - Sets the intermission between delivery of each mail.


(1) - Create bulk emails in account's draft folder for examination.
(2) - Start mail delivery.
"""

# --- German
    def SetDe(self):
        '''set German'''

        # labels
        self.Lang_From.set("Von:")
        self.Lang_Subject.set("Betreff:")
        self.Lang_Recipient.set("Empfängerdatei:")
        self.Lang_uId.set("ID-Spalte:")
        self.Lang_Mail.set("E-Mail-Spalte:")
        self.Lang_Html.set("HTML-Datei:")
        self.Lang_Text.set("Text-Datei:")
        self.Lang_Attach.set("Anhänge hinzufügen:")
        self.Lang_Format.set("Format:")
        self.Lang_Receipt.set("Empfangsbestätigung anfordern:")
        self.Lang_Interval.set("Versandintervall in Sek.:")
        self.Lang_Address.set("Serienmail-Daten")
        self.Lang_Content.set("Mailinhalt")
        self.Lang_Parse.set("" "(1) E-Mail-Entwürfe anlegen")
        self.Lang_Send.set("(2) Entwürfe versenden" )
        self.Lang_Help.set("Hinweise")
        self.Lang_Log.set("Log-Ordner:")
        self.Lang_File.set("Datei:")
        self.Lang_Folder.set("Ordner:")
        self.Lang_TError = "Das gewählt Versandformat erfordert die Auswahl einer Text-Datei."
        self.Lang_HError = "Das gewählt Versandformat erfordert die Auswahl einer HTML-Datei."
        self.Lang_THError = "Das gewählt Versandformat erfordert die Auswahl einer Text- und einer HTML-Datei."
        self.Lang_FormError = "Bitte überprüfen Sie Ihre Eingabe.\nEs wurden nicht alle erforderlichen Felder befüllt."
        self.hint = """
           'Von' - Versenderadresse
       'Betreff' - E-Mail-Betreff

            Serienmail-Daten
'Empfängerdatei' - XLSX-Datei mit Serienmaildaten
                   Daten befinden sich im ersten Tabellenblatt.
     'ID-Spalte' - Spaltenbezeichnung (Zeile 1) in der XLSX mit der ID des Eintrags.
                   Die ID ist die Grundlage der automatischen Auswahl der Anhänge.
 'E-Mail-Spalte' - Spaltenbezeichnung der Emppfänger-E-Mail-Spalten
Empfangsbestätigung anfordern - Aktiviert die Forderung einer Empfangsbestätigung.

            Anhänge
Mit dem "+" Symbol kann eine arbiträre Anzahl an Anhängen beigefügt werden.
       'Datei'  - Einzeldatei die JEDER Mail beigefügt wird.
       'Ordner' - Sammlung von Dateien, die anhand der ID der JEWEILIGEN Mail beigefügt werden.

            Mailinhalt
E-Mails können als Text, Html oder beides versandt werden. Dafür müssen die entsprechenden
Dateien vorliegen. Diese können z.B. in Word erstellt werden. Das Programm führt einen
Textersatz vergleichbar mit einem Serienbrief im Mailkörper anhand der Daten in der
Empfägnerdatei durch. Textersatz im Mailtext wird als {Spaltenname} markiert.

    'Log-Ordner' - Ordner in welchem die Log-Datei gespeichert wird.
Versandintervall - Legt das Intervall zwischem dem Senden zweier Mails fest.

(1) - Zuerst werden die Mail im Ordner Entwürfe angelegt und können dort geprüft werden.
(2) - Startet den Versand der Mails.
"""


# --- create window
root = Tk()
my_gui = CoreGui(root)
root.grid_columnconfigure(1, weight=5)
root.mainloop()
