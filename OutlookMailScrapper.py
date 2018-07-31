# -*- coding: utf-8 -*-
"""
Created on Mon Jul  9 16:45:59 2018

@author: shardsin
"""

import win32com.client   # Will help us to access the Outlook Application
import re                # Regular Expression Check
import pandas as pd      # For creating the dataset
import matplotlib.pyplot as plt    # Used for creating graph
import numpy as np       # Will create arrays to help create graph data
        
class OutlookMail:
    
    """ 
    Creating a list which will contain set of Dictionaries containing
    Sender's Name, Subject and Time of the mail
    """
    def __init__(self):
        
        self.items = []
        self.mailDict = dict()
        self.senders = set()

    """ 
    encodeValue will be used to convert the non-string values fetched from the 
    mail to String
    """
    def encodeValue(self,s):
        if isinstance(s, str):
            return s
        else:
            return str(s)

    """
    This is the main method
    It connects to the Outlook Application using the win32com.client package
    Gets access to the Mails using the MAPI (Mail API)
    Reads the mail in the Inbox Folder
    Extract the Subject, Sender and Time from every mail in the Inbox folder
    Filter for a particular Subject
    Creates the list containting Dictionaries
    """
    
    def extractMail(self):
    
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
                "MAPI")
        inbox    = outlook.GetDefaultFolder(6)  # "6" refers to the inbox
        messages = inbox.Items
        message  = messages.GetLast()
    
    
        while message:
    
            try:
                d = dict()
                d['Subject'] = self.encodeValue(message.Subject)
                d['Received Time'] = self.encodeValue(message.ReceivedTime)
                if re.search("Adhoc: .", d['Subject'], re.IGNORECASE):
                    d['Sender']  = self.encodeValue(message.Sender)
                    self.items.append(d)
    
            except Exception:
                d['Sender'] = "Meeting Invite"
                self.items.append(d) 
            
            message = messages.GetPrevious()    

    """
    writeCSV method uses pandas library to create a Data Frame from the
    existing list and creates a CSV file
    """
    
    def writeCSV(self):
        df = pd.DataFrame(self.items)
        df.to_csv("Mails.csv", index=False)

    def readCSV(self):
        dataFrame = pd.read_csv("Mails.csv")
        sendersNameList = list (str(i[0]) for i in dataFrame.iloc[:,1:-1].
                                values)
        self.senders = set(sendersNameList)
        for i in self.senders:
            key = i.split(" ")[0]
            mailNums = sendersNameList.count(i)
            self.mailDict[key] = mailNums

    def plotGraph(self):

        fig, ax = plt.subplots()
        y = np.arange(len(self.senders))
        ax.barh(y, list(self.mailDict.values()), align="center", color="blue", 
                ecolor="black")
        ax.set_yticks(y)
        ax.set_yticklabels(list(self.mailDict.keys()))
        ax.invert_yaxis()
        plt.xlabel("Number of Mails")
        plt.show()
        
if __name__ == "__main__":
    mail = OutlookMail()
    mail.extractMail()
    mail.writeCSV()
    mail.readCSV()
    mail.plotGraph()