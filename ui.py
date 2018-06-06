# -*- coding: utf-8 -*-

#!/usr/bin/env python3

import tkinter as tk
from tkinter import Tk, Label, Button, StringVar, Entry, filedialog, W, E, N, S
from tkinter import Radiobutton, Checkbutton, IntVar
from tkinter import messagebox as mbox
from time import sleep

from os import listdir
#import pandas as pd
from pandas import read_excel
#import win32com.client


class CoreGui(tk.Frame):

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)

        self.attachFields = [] # stores UI attachment fields
        self.attachments = {} # stores attachment references
        self.mailForm = IntVar()
        self.recConf = IntVar()

        self.master = master
        master.title("MAPI Massmailer")

        self.label = Label(master, text="Programm zum Versenden von Massen-E-Mails")

        self.senderMailLab = Label(master, text="Von:")
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
        self.mailForm.set(3)
        self.mailFormat1 = Radiobutton(master, text="Text",
                                       variable=self.mailForm, value=1)
        self.mailFormat2 = Radiobutton(master, text="Html",
                                       variable=self.mailForm, value=2)
        self.mailFormat3 = Radiobutton(master, text="Text & Html",
                                       variable=self.mailForm, value=3)

        self.receiptConfirm = Checkbutton(master, text="mit Empfangsbestätigung",
                                          variable=self.recConf)

        self.space = Label(master,text=" ")
        self.space1 = Label(master,text=" ")
        self.space2 = Label(master,text=" ")
        self.address = Label(master, text="Adressdatei")
        self.mailtext = Label(master, text="Mailinhalt")

        self.v = IntVar()
        self.pauseLab = Label(master, text="Versandintervall in Sek.:")
        self.pause = Entry(master, text=self.v, width=4)
        self.v.set(10)

        self.send = Button(master, text="Senden", command=self.ParseMail)
        self.help = Button(master, text="Hilfe", command=self.onInfo)

        # layout the program
        self.send.grid(row=0, column=0, sticky=W)
        self.help.grid(row=0, column=1, sticky=E)
        self.senderMailLab.grid(row=1, column=0, sticky=E)
        self.senderMail.grid(row=1, column=1)
        #self.senderAliasLab.grid(row=2, column=0, sticky=E)
        #self.senderAlias.grid(row=2, column=1)
        self.subjectLab.grid(row=2, column=0, sticky=E)
        self.subject.grid(row=2, column=1)
        self.space.grid(row=3, columnspan=2)
        self.address.grid(row=4, columnspan=2)
        self.recipientFileSel.grid(row=5, column=0, sticky=E)
        self.recipientFile.grid(row=5, column=1)
        self.uniqueIdLab.grid(row=6, column=0, sticky=E)
        self.uniqueId.grid(row=6, column=1, sticky=W)
        self.mailIdLab.grid(row=7, column=0, sticky=E)
        self.mailId.grid(row=7, column=1, sticky=W)
        self.space1.grid(row=8, columnspan=2)
        self.mailtext.grid(row=9, columnspan=2)
        self.mailFormatLab.grid(row=10, column=0)
        self.mailFormat1.grid(row=10, column=1, sticky=W)
        self.mailFormat2.grid(row=10, column=1)
        self.mailFormat3.grid(row=10, column=1, sticky=E)
        self.mailBodyHtmlSel.grid(row=11, column=0, sticky=E)
        self.mailBodyHtml.grid(row=11, column=1)
        self.mailBodySel.grid(row=12, column=0, sticky=E)
        self.mailBody.grid(row=12, column=1)
        self.space2.grid(row=13, columnspan=2)
        self.pauseLab.grid(row=14, column=0, sticky=E)
        self.pause.grid(row=14, column=1, sticky=W)
        self.receiptConfirm.grid(row=14, column=1, sticky=E)
        self.AttachmentsLab.grid(row=15, columnspan=2)
        self.addAttachment.grid(row=15, column=1, sticky=E)

        self.grid(columnspan=2, sticky="NEWS")

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
        mbox.showinfo("Nutzung", "Empfängerdatei: Ist eine *.xlsx Datei mit Adressdaten und Textersatzdaten.\nE-Mails können in HTML und Txt-Format versandt werden.\nTextersatzfelder werden markiert mit {Spalte} aus der Empfängerdatei.")

    def ParseMail(self):
        '''collate all necessary data for single email and send'''

        senderMail = self.senderMail.get()
        senderAlias = self.senderAlias.get()
        subject = self.subject.get()
        recipient = self.recipientDf
        uniqueId = self.uniqueId.get()
        mailBodyHtml = self.mailBodyHtmlRaw
        mailBody = self.mailBodyRaw
        attach = self.attachments

        # create outlook session
        mapiSes = win32com.client.Dispatch("Mapi.Session")
        appliOut = win32com.client.Dispatch("Outlook.Application")
        mapiSes.Logon("Outlook2010")

        for i in range(0, len(recipient.index)):
            replacement = recipient.iloc[[i]].to_dict('records')
            # print(replacement)
            if self.mailForm == 1:
                bodyFormatTxt = mailBody.format(**replacement[0])
            elif self.mailForm == 2:
                bodyFormatHtml = mailBodyHtml.format(**replacement[0])
            else:
                bodyFormatTxt = mailBody.format(**replacement[0])
                bodyFormatHtml = mailBody.format(**replacement[0])
            # print(bodyFormat)

            # create message
            Msg = appliOut.CreateItem(0)
            Msg.From = senderMail
            Msg.Subject = subject
            Msg.To = replacement[uniqueId]

            if self.mailForm == 1:
                Msg.Body = bodyFormatTxt
            elif self.mailForm == 2:
                Msg.HTMLBody = bodyFormatHtml
            else:
                Msg.Body = bodyFormatTxt
                Msg.HTMLBody = bodyFormatHtml

            # Read receipt
            if self.recConf == 1:
                Msg.ReadReceipt = True

            # add attachments
            for j in range(0, len(attach)):
                # single file
                if type(attach[j]) is str:
                    Msg.Attachments.Add(attach[j])
                # multiple files from folder
                elif type(attach[j]) is list:
                    matches = [s for s in attach[j][1] if replacement['uniqueId'] in s]
                    if len(matches) > 0:
                        for el in matches:
                            addAttach = attach[j]+'/'+el
                            Msg.Attachments.Add(addAttach)

            # send message
            Msg.Send()
            sleep(self.pause.get())


root = Tk()
my_gui = CoreGui(root)
root.mainloop()
