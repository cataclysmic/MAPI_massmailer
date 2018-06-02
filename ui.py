# -*- coding: utf-8 -*-

#!/usr/bin/env python3

import tkinter as tk
from tkinter import Tk, Label, Button, StringVar, Entry, filedialog, W, E, N, S


class CoreGui(tk.Frame):

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)

        self.attachFields = []

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
        self.recipientFile = Entry(master, width=40)

        self.uniqueIdLab = Label(master, text="Kennfeld:")
        self.uniqueId = Entry(master, width=40)

        self.mailBodySel = Button(master, text="Mail Body:",
                                  command=self.LoadBody)
        self.mailBody = Entry(master, width=40)

        self.AttachmentsLab = Label(master, text="Anhänge:")
        self.addAttachment = Button(master, text="+",
                                    command=self.AddAttachmentField)

        self.send = Button(master, text="Senden", command=self.SendMail)


        self.send.grid(row=0, column=1, sticky=E)
        self.senderMailLab.grid(row=1, column=0)
        self.senderMail.grid(row=1, column=1)
        self.senderAliasLab.grid(row=2, column=0)
        self.senderAlias.grid(row=2, column=1)
        self.subjectLab.grid(row=3, column=0)
        self.subject.grid(row=3, column=1)
        self.recipientFileSel.grid(row=4, column=0)
        self.recipientFile.grid(row=4, column=1)
        self.uniqueIdLab.grid(row=5,column=0)
        self.uniqueId.grid(row=5,column=1)
        self.mailBodySel.grid(row=6, column=0)
        self.mailBody.grid(row=6, column=1)
        self.AttachmentsLab.grid(row=7, columnspan=2)
        self.addAttachment.grid(row=7, column=1, sticky=E)

        self.grid(columnspan=2,sticky="NEWS")

    def LoadRecipient(self):
        '''Load spreadsheet file with recipients.'''

        filename = filedialog.askopenfilename(
            filetypes=(("XLSX", "*.xlsx"),
                       ("CSV", "*.csv")))

        return(filename)

    def LoadBody(self):
        '''Load html file contain mail body'''

        filename = filedialog.askopenfilename(
            filetypes=(("HTML", "*.html"),
                       ("HTML", "*.htm")))

        return(filename)

    def AddAttachmentField(self):
        '''add attachment via + '''

        #super(CoreGui, self).__init__()

        n = len(self.attachFields)
        i = n
        if i > 0:
            i -= i
        self.attachFields.append({})
        #self.attachFields[n]['label'] = Label(self, text="("+str(n+1)+")")
        #self.attachFields[n]['label'].grid(row=7+n, column=1, sticky=E)

        self.attachFields[i]['folderBut'] = Button(self, text="Ordner")
        self.attachFields[i]['fileBut'] = Button(self, text="Datei")
        self.attachFields[i]['field'] = Entry(self, width=40)

        print(self.attachFields[i])

        self.attachFields[i]['fileBut'].grid(row=n, column=0)
        self.attachFields[i]['folderBut'].grid(row=n, column=1)
        self.attachFields[i]['field'].grid(row=n, column=2)


    def SendMail(self):
        '''Execute batch mailing'''


root = Tk()
my_gui = CoreGui(root)
root.mainloop()
