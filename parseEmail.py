import email
from email import parser
import Tkinter
import tkFileDialog
import os
import pathlib2
import xlsxwriter
import tkMessageBox
from bs4 import BeautifulSoup
import poplib

root = Tkinter.Tk()
root.withdraw()

# Initialize list that will hold instance,type,and date
list = []

# Connect to the specified gmail account
SERVER = "pop.gmail.com"
EMAIL = ""
PASS = ""

print "Connecting to gmail server..."
server = poplib.POP3_SSL(SERVER)

print "Logging in..."
server.user(EMAIL)
server.pass_(PASS)

print "You are now logged in as " + EMAIL
messages = [server.retr(i) for i in range(1, len(server.list()[1]) + 1)]

# If there are any new emails
if messages:
    # Concat message pieces:
    messages = ["\n".join(mssg[1]) for mssg in messages]
    #Parse message intom an email object:
    messages = [parser.Parser().parsestr(mssg) for mssg in messages]
    for theEmail in messages:
        #Gets the instance and date
        subject = theEmail['subject']
        instance = subject.split("-",1)[1].strip()
        date = theEmail['date']
        payloadList = []
        # Get the body of the email
        # TO-DO
        for payload in theEmail.get_payload():
            payloadList.append(payload.get_payload())
        soup = BeautifulSoup(payloadList[1],"lxml")
        # Get the type of email
        if "ALERT" in subject:
            type = "Alert"
        elif "RESOLVED" in subject:
            type = "Resolved"
        else:
            print ("This email is not an alert or a resolved")

        # Create tuple with information and add to list
        tmp = (instance,type,date)
        list.append(tmp)
    # If the list is not empty, create excel sheet
    if list:
        # Creates excel workbook and sheet
        workbook = xlsxwriter.Workbook('PATInformation.xlsx')
        worksheet = workbook.add_worksheet()
        # Creates bold format
        bold = workbook.add_format({'bold': True})
        # Adds headers for the columns and sets width
        worksheet.write('A1','Instance',bold)
        worksheet.set_column('A:A',55)
        worksheet.write('B1','Type',bold)
        worksheet.set_column('C:C',30)
        worksheet.write('C1','Date',bold)

        #worksheet.write('D1','Body')
        row = 1
        col = 0
        for x in list:
            # Instance
            worksheet.write(row,col,x[0])
            # Type
            worksheet.write(row,col + 1,x[1])
            # Date
            worksheet.write(row,col + 2,x[2])
            # Body
            #worksheet.write(row,col + 3,x[3])
            row = row + 1
        tkMessageBox.showinfo("PAT-Automation","Excel sheet has been created successfully!")
        workbook.close()
else:
    print "There are no new emails"
server.quit()
