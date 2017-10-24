import email
import Tkinter
import tkFileDialog
import os
import pathlib2
import xlsxwriter

root = Tkinter.Tk()
root.withdraw()

# Creates excel workbook and sheet
workbook = xlsxwriter.Workbook('PATInformation.xlsx')
worksheet = workbook.add_worksheet()
# Adds headers for the columns
worksheet.write('A1','Instances',bold)
worksheet.write('B1','Dates',bold)

directory = tkFileDialog.askdirectory(initialdir="~/")
if directory:
    for filename in os.listdir(directory):
        path = directory + "/" + filename
        if (pathlib2.Path(path).suffix == ".eml" or pathlib2.Path(path).suffix == ".msg"):
            file = open(path)
            theEmail = email.message_from_file(file)
            file.close()
            subject = theEmail['subject']
            instance = subject.split("-",1)[1].strip()
            print instance
            worksheet.write('A1',instance)
            date = theEmail['date']
            print date
            worksheet.write('A2',date)
            # if "ALERT" in subject:
            #     print ("ALERT")
            # elif "RESOLVED" in subject:
            #     print ("RESOLVED")
            # else:
            #     print ("This email is not an alert or a resolved")
        else:
            print ("The file " + filename + " is not a .eml or .msg so skipping")
workbook.close()
