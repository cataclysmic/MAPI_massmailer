# -*- coding: utf-8 -*-

#!/usr/bin/env python3

import tkinter as tk
from tkinter import Tk, Label, Button, StringVar, Entry, filedialog, W, E, N, S
from os import listdir
import pandas as pd


class CoreGui(tk.Frame):

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)

        self.attachFields = [] # stores UI attachment fields
        self.attachments = {} # stores attachment references

        self.master = master
        master.title("MAPI Massmailer")

        self.label = Label(master, text="Programm zum Versenden von Massen-E-Mails")

        self.senderMailLab = Label(master, text="Versand E-Mail:")
        self.senderMail = Entry(master, width=40)

        self.senderAliasLab = Label(master, text="Sender Alias:")
        self.senderAlias = Entry(master, width=40)

        self.subjectLab = Label(master, text="Betreff:")
        self.subject = Entry(master, width=40)

        self.recipientFileSel = Button(master, text="Empfängerdatei:",
                                       command=self.LoadRecipient)
        self.recipientFileStr = StringVar()
        self.recipientFile = Entry(master, width=40, state='readonly',
                                   textvariable=self.recipientFileStr)

        self.uniqueIdLab = Label(master, text="ID Spalte:")
        self.uniqueId = Entry(master, width=40)

        self.mailBodySel = Button(master, text="Mail Body:",
                                  command=self.LoadBody)
        self.mailBodyStr = StringVar()
        self.mailBody = Entry(master, width=40, state='readonly',
                              textvariable=self.mailBodyStr)

        self.AttachmentsLab = Label(master, text="Anhänge:")
        self.addAttachment = Button(master, text="+",
                                    command=self.AddAttachmentField)

        self.send = Button(master, text="Senden", command=self.ParseMail)


        self.send.grid(row=0, column=1, sticky=E)
        self.senderMailLab.grid(row=1, column=0, sticky=E)
        self.senderMail.grid(row=1, column=1)
        self.senderAliasLab.grid(row=2, column=0, sticky=E)
        self.senderAlias.grid(row=2, column=1)
        self.subjectLab.grid(row=3, column=0, sticky=E)
        self.subject.grid(row=3, column=1)
        self.recipientFileSel.grid(row=4, column=0, sticky=E)
        self.recipientFile.grid(row=4, column=1)
        self.uniqueIdLab.grid(row=5,column=0, sticky=E)
        self.uniqueId.grid(row=5,column=1, sticky=W)
        self.mailBodySel.grid(row=6, column=0, sticky=E)
        self.mailBody.grid(row=6, column=1)
        self.AttachmentsLab.grid(row=7, columnspan=2)
        self.addAttachment.grid(row=7, column=1, sticky=E)

        self.grid(columnspan=2,sticky="NEWS")

    def LoadRecipient(self):
        '''Load spreadsheet file with recipients.'''

        filename = filedialog.askopenfilename(
            filetypes=(("XLSX", "*.xlsx"),
                       ("Alle Dateien", "*.*")))

        self.recipientFileStr.set(filename)

        self.recipientDf = pd.read_excel(filename)


    def LoadBody(self):
        '''Load html file contain mail body'''

        filename = filedialog.askopenfilename(
            filetypes=(("HTML", "*.html"),
                       ("HTML", "*.htm")))

        self.mailBodyStr.set(filename)

        file = open(filename,'r')
        self.mailBodyRaw = file.read()

    def AddAttachmentField(self):
        '''add attachment via + '''

        n = len(self.attachFields)
#        print(n)
        self.attachFields.append({})

        self.attachFields[n]['label'] = Label(self, text="("+str(n+1)+")")

        self.attachFields[n]['folderBut'] = Button(self, text="Ordner",
                                                 command=lambda: self.LoadAttachFolder(n))
        self.attachFields[n]['fileBut'] = Button(self, text="Datei",
                                                 command=lambda: self.LoadAttachFile(n))
        self.attachFields[n]['stringVar'] = StringVar()
        self.attachFields[n]['field'] = Entry(self, width=37, state="readonly",
                                              textvariable=self.attachFields[n]['stringVar'])

#        print(self.attachFields[n])

        self.attachFields[n]['fileBut'].grid(row=n, column=0)
        self.attachFields[n]['folderBut'].grid(row=n, column=1)
        self.attachFields[n]['field'].grid(row=n, column=2)
        self.attachFields[n]['label'].grid(row=n, column=3)

    def LoadAttachFile(self, idNr):
        '''load attachment file path'''

        filename = filedialog.askopenfilename(
            filetypes=(("Alle Dateien", "*.*"),
                       ("PDF", "*.pdf")))

        self.attachFields[idNr]['stringVar'].set(filename)
        self.attachments[idNr] = filename

        print(idNr)
        print(self.attachments)

    def LoadAttachFolder(self,idNr):
        '''load folder path and list'''

        foldername = filedialog.askdirectory()
        self.attachFields[idNr]['stringVar'].set(foldername)

        fileList = listdir(foldername)
        self.attachments[idNr] = [foldername, fileList]

        print(idNr)
        print(self.attachments)

    def ParseMail(self):
        '''collate all necessary data for sing email'''

        senderMail = self.senderMail.get()
        senderAlias = self.senderAlias.get()
        subject = self.subject.get()
        recipient = self.recipientDf
        uniqueId = self.uniqueId.get()
        mailBody = self.mailBodyRaw
        attach = self.attachments

        for i in range(0,len(recipient.index)):
            replacement = recipient.iloc[[i]].to_dict('records')
            print(replacement)
            bodyFormat = mailBody.format(**replacement[0])

            print(bodyFormat)




    def SendMail(self):
        '''Execute batch mailing'''


root = Tk()
my_gui = CoreGui(root)
root.mainloop()
