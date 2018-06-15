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

        self.master = master
        master.title("MAPI Massmailer")

        self.label = Label(master, text="Programm zum Versenden von Massen-E-Mails")

        self.senderMailLab = Label(master, text="Von:")
        self.senderMail = Entry(master, width=40)

        self.sv = StringVar()
        self.storeFolderLab = Label(master, text="Speicherordner:")
        self.storeFolder = Entry(master, width=40, text=self.sv)
        self.sv.set("Entwürfe")

        self.subjectLab = Label(master, text="Betreff:")
        self.subject = Entry(master, width=40)

        self.recipientFileSel = Button(master, text="Empfängerdatei:",
                                       command=self.LoadRecipient)
        self.recipientFileStr = StringVar()
        self.recipientFile = Entry(master, width=40, state='readonly',
                                   textvariable=self.recipientFileStr)

        self.uniqueIdLab = Label(master, text="ID Spalte (z.B. AGS):")
        self.uniqueId = Entry(master, width=40)

        self.mailIdLab = Label(master, text="E-Mail Spalte (z.B. EMAIL):")
        self.mailId = Entry(master, width=40)

        self.mailBodyHtmlSel = Button(master, text="HTML-Datei:",
                                      command=self.LoadBodyHtml)
        self.mailBodyHtmlStr = StringVar()
        self.mailBodyHtml = Entry(master, width=40, state='readonly',
                                  textvariable=self.mailBodyHtmlStr)

        self.mailBodySel = Button(master, text="Text-Datei:",
                                  command=self.LoadBodyText)
        self.mailBodyStr = StringVar()
        self.mailBody = Entry(master, width=40, state='readonly',
                              textvariable=self.mailBodyStr)

        self.AttachmentsLab = Label(master, text="Anhänge:")
        self.addAttachment = Button(master, text="+",
                                    command=self.AddAttachmentField)

        self.mailFormatLab = Label(master, text="Format:")
        self.mailForm = IntVar()
        self.mailForm.set(3)
        self.mailFormat1 = Radiobutton(master, text="Text",
                                       variable=self.mailForm, value=1)
        self.mailFormat2 = Radiobutton(master, text="Html",
                                       variable=self.mailForm, value=2)
        self.mailFormat3 = Radiobutton(master, text="Text & Html",
                                       variable=self.mailForm, value=3)

        self.receiptConfirm = Checkbutton(master, text="Empfangsbestätigung anfordern.",
                                          variable=self.recConf)

        self.space = Label(master,text=" ")
        self.space1 = Label(master,text=" ")
        self.space2 = Label(master,text=" ")
        self.address = Label(master, text="Adressdatei")
        self.mailtext = Label(master, text="Mailinhalt")

        self.v = IntVar()
        self.pauseLab = Label(master, text="Versandintervall in Sek.:")
        self.pause = Entry(master, text=self.v, width=4)
        self.v.set(5)

        self.parse = Button(master, text="(1) Mail-Entwürfe anlegen", command=self.ParseMail)
        self.send = Button(master, text="(2) Mails versenden", command=self.SendMail)
        self.help = Button(master, text="Hinweise", command=self.onInfo)

        self.logLocationStr = StringVar()
        self.logLoc = Button(master, text="Log-Ordner:", command=self.LogFolder)
        self.logLocation = Entry(master, width=40, state='readonly',
                                 textvariable=self.logLocationStr)

        # layout the program
        self.help.grid(row=0, column=1, sticky=E)
        self.senderMailLab.grid(row=1, column=0, sticky=E)
        self.senderMail.grid(row=1, column=1)
        self.storeFolderLab.grid(row=2, column=0, sticky=E)
        self.storeFolder.grid(row=2, column=1)
        self.subjectLab.grid(row=3, column=0, sticky=E)
        self.subject.grid(row=3, column=1)
        self.space.grid(row=4, columnspan=2)
        self.address.grid(row=5, columnspan=2)
        self.recipientFileSel.grid(row=6, column=0, sticky=E)
        self.recipientFile.grid(row=6, column=1)
        self.uniqueIdLab.grid(row=7, column=0, sticky=E)
        self.uniqueId.grid(row=7, column=1, sticky=W)
        self.mailIdLab.grid(row=8, column=0, sticky=E)
        self.mailId.grid(row=8, column=1, sticky=W)
        self.space1.grid(row=9, columnspan=2)
        self.mailtext.grid(row=10, columnspan=2)
        self.mailFormatLab.grid(row=11, column=0)
        self.mailFormat1.grid(row=11, column=1, sticky=W)
        self.mailFormat2.grid(row=11, column=1)
        self.mailFormat3.grid(row=11, column=1, sticky=E)
        self.mailBodyHtmlSel.grid(row=12, column=0, sticky=E)
        self.mailBodyHtml.grid(row=12, column=1)
        self.mailBodySel.grid(row=13, column=0, sticky=E)
        self.mailBody.grid(row=13, column=1)
        self.receiptConfirm.grid(row=15, column=1, sticky=W)
        self.AttachmentsLab.grid(row=16, columnspan=2)
        self.addAttachment.grid(row=16, column=1, sticky=E)

        self.grid(columnspan=2, sticky="NEWS")

        self.parse.grid(columnspan=2, sticky="EW")
        self.logLoc.grid(row=20, column=0, sticky=E)
        self.logLocation.grid(row=20, column=1, sticky=E)
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
                       ("Alle Dateien", "*.*")))

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

    def LoadBodyHtml(self):
        '''Load html file contain mail body'''

        filename = filedialog.askopenfilename(
            filetypes=(("HTML", "*.html"),
                       ("HTML", "*.htm")))

        self.mailBodyHtmlStr.set(filename)

        file = open(filename, 'r')
        self.mailBodyHtmlRaw = file.read()

    def AddAttachmentField(self):
        '''add attachment via + '''

        n = len(self.attachFields)
#        print(n)
        self.attachFields.append({})

        self.attachFields[n]['label'] = Label(self, text="("+str(n+1)+")")
        self.attachFields[n]['space'] = Label(self, text="   ")

        self.attachFields[n]['folderBut'] = Button(self, text="Ordner",
                                                 command=lambda: self.LoadAttachFolder(n))
        self.attachFields[n]['fileBut'] = Button(self, text="Datei",
                                                 command=lambda: self.LoadAttachFile(n))
        self.attachFields[n]['stringVar'] = StringVar()
        self.attachFields[n]['field'] = Entry(self, width=43, state="readonly",
                                              textvariable=self.attachFields[n]['stringVar'])

#        print(self.attachFields[n])

        self.attachFields[n]['fileBut'].grid(row=n, column=0)
        self.attachFields[n]['folderBut'].grid(row=n, column=1)
        self.attachFields[n]['space'].grid(row=n, column=2)
        self.attachFields[n]['field'].grid(row=n, column=3, sticky="E")
        self.attachFields[n]['label'].grid(row=n, column=4, sticky="E")

    def LoadAttachFile(self, idNr):
        '''load attachment file path'''

        filename = filedialog.askopenfilename(
            filetypes=(("Alle Dateien", "*.*"),
                       ("PDF", "*.pdf")))

        self.attachFields[idNr]['stringVar'].set(filename)
        self.attachments[idNr] = filename

        print(idNr)
        print(self.attachments)

    def LoadAttachFolder(self, idNr):
        '''load folder path and list'''

        foldername = filedialog.askdirectory()
        self.attachFields[idNr]['stringVar'].set(foldername)

        fileList = listdir(foldername)
        self.attachments[idNr] = [foldername, fileList]

        print(idNr)
        print(self.attachments)

    def onInfo(self):
        '''helpful message '''
        mbox.showinfo("Nutzung",
"""
           'Von' - Versenderadresse
'Speicherordner' - Zielordner für das Anlegen der E-Mails
       'Betreff' - E-Mail-Betreff

            Adressdatei
'Empfängerdatei' - XLSX-Datei mit Serienmaildaten
                   Daten befinden sich im ersten Tabellenblatt.
     'ID Spalte' - Spaltenbezeichnung (Zeile 1) in der XLSX mit der ID des Eintrags.
                   Die ID ist die Grundlage der automatischen Auswahl der Anhänge.
 'E-Mail Spalte' - Spaltenbezeichnung der Emppfänger-E-Mail-Spalten

            Mailinhalt
E-Mail können als Text, Html oder beides versandt werden. Dafür müssen die entsprechenden
Dateien vorliegen. Diese können z.B. in Word erstellt werden. Das Programm führt einen
Textersatz vergleichbar mit einem Serienbrief im Mailkörper anhand der Daten in der
Empfägnerdatei durch. Textersatz im Mailtext wird als {Spaltenname} markiert.

Versandintervall - Legt das Intervall zwischem dem Senden zweier Mails fest.
mit Empfangsbestätigung - Aktiviert die Forderung einer Empfangsbestätigung.


            Anhänge
Mit dem "+" Symbol kann eine arbiträre Anzahl an Anhängen beigefügt werden.
       'Datei'  - Einzeldatei die JEDER Mail beigefügt wird.
       'Ordner' - Sammlung von Dateien, die anhand der ID der JEWEILIGEN Mail beigefügt werden.


(1) - Zuerst werden die Mail im Ordner Entwürfe angelegt und können dort geprüft werden.
(2) - Startet den Versand der Mails.
""")


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
            mbox.showwarning("Fehler",
                             "Bitte überprüfen Sie Ihre Eingabe.\nEs wurden nicht alle erforderlichen Felder befüllt.")
            return()


        if int(self.mailForm.get()) == 1 and self.mailBodyRaw.strip() == "":
            mbox.showinfo("Das gewählt Versandformat erfordert die Auswahl einer Text-Datei.")
            return("NO")
        elif int(self.mailForm.get()) == 2 and self.mailBodyHtmlRaw.strip() == "":
            mbox.showinfo("Das gewählt Versandformat erfordert die Auswahl einer HTML-Datei.")
            return("NO")
        elif int(self.mailForm.get()) == 3 and (self.mailBodyHtmlRaw.strip() == "" or self.mailBodyRaw.strip() == ""):
            mbox.showinfo("Das gewählt Versandformat erfordert die Auswahl einer Text- und einer HTML-Datei.")
            return("NO")

        mailForm = int(self.mailForm.get())
        print(uniqueId)


        OutBox = self.GetOutbox()
        #appliOut = Dispatch("Outlook.Application").GetNamespace("MAPI")

        #OutBox = appliOut.GetDefaultFolder(5)

        for i in range(0, len(recipient.index)):
            replacement = recipient.iloc[[i]].to_dict('records')
            print(replacement)
            if mailForm == 1:
                mailBody = self.mailBodyRaw
                bodyFormatTxt = mailBody.format(**replacement[0])
            elif mailForm == 2:
                mailBodyHtml = self.mailBodyHtmlRaw
                mailBodyHtml = mailBodyHtml.split("<body ")[1].split("\n",1)[1].split("</body>")[0]
                bodyFormatHtml = mailBodyHtml.format(**replacement[0])
            else:
                mailBody = self.mailBodyRaw
                mailBodyHtml = self.mailBodyHtmlRaw
                mailBodyHtml = mailBodyHtml.split("<body ")[1].split("\n",1)[1].split("</body>")[0]
                bodyFormatTxt = mailBody.format(**replacement[0])
                bodyFormatHtml = mailBodyHtml.format(**replacement[0])

            # create message
            Msg = OutBox.Items.Add(0)
            #Msg = appliOut.CreateItem(0x0)
            Msg.Subject = subject
            Msg.To = replacement[0][mailId]
            Msg.SentOnBehalfOfName = senderMail

            if mailForm == 1:
                Msg.Body = bodyFormatTxt
            elif mailForm == 2:
                Msg.HTMLBody = bodyFormatHtml
            else:
                Msg.Body = bodyFormatTxt
                Msg.HTMLBody = bodyFormatHtml

            # Read receipt
            if int(self.recConf.get()) == 1:
                Msg.ReadReceiptRequested = True

            # add attachments
            for j in range(0, len(attach)):
                # single file
                if type(attach[j]) is str:
                    Msg.Attachments.Add(attach[j])
                # multiple files from folder
                elif type(attach[j]) is list:
                    matches = [s for s in attach[j][1] if str(replacement[0][uniqueId]) in s]
                    if len(matches) > 0:
                        for el in matches:
                            addAttach = attach[j][0]+'/'+el
                            Msg.Attachments.Add(addAttach)

            # send message
            #Msg.display()
            Msg.Save()

    def SendMail(self):
        '''iterate through folder and send mails'''

        logPath = self.logLocation.get()+"/"+"massmail_log_"+strftime("%Y%m%d-%H%M%S")

        file = open(logPath+".log", "a")
        file.write(self.subject.get()+"\n\n")
        OutBox = self.GetOutbox()
        messages = OutBox.Items
        for i in range(0,len(messages)):
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
        storeFolder = self.storeFolder.get()

        appliOut = Dispatch("Outlook.Application").GetNamespace("MAPI")
        for i in range(1,20):
            accounts = appliOut.Folders(i)
            if str(accounts) == fromMail:
                for j in range(1,20):
                    boxes = appliOut.Folders(i).Folders(j)
                    if str(boxes) == storeFolder:
                        break
                break
        print(boxes)
        return(boxes)


root = Tk()
my_gui = CoreGui(root)
root.mainloop()
