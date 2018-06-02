# -*- coding: utf-8 -*-

#!/usr/bin/env python3

from tkinter import Tk, Label, Button, StringVar, Entry, filedialog


class CoreGui:

    def __init__(self, master):
        self.master = master
        master.title("MAPI Massmailer")

        self.senderMailLab = StringVar()
        self.senderMailLab.set("Versand-E-Mail:")
        self.senderMail = Entry(master)

        self.senderAliasLab =StringVar()
        self.senderAliasLab.set("Sender Alias:")
        self.senderAlias = Entry(master)

        self.subjectLab = StringVar()
        self.subjectLab.set("Betreff:")
        self.subject = Entry(master)

        self.recipientFile = Button(master, text="Empfängerdatei:",
                                    command=self.LoadRecipient)
        self.recipientFileLab = StringVar()
        self.recipientFileLab.set("LEER")

        self.mailBodyLab = StringVar()
        self.mailBodyLab.set("Inhalt:")
        self.mailBody = Button(master, text="Inhaltsdatei:",
                               command=self.LoadBody)

        self.nrAttachmentsLab = StringVar()
        self.nrAttachmentsLab.set("Anzahl der Anhänge:")
        self.nrAttachments = Entry(master)

