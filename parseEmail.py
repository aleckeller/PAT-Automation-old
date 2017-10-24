import email
import Tkinter
import tkFileDialog
import os
import pathlib2
import xlsxwriter
import tkMessageBox

root = Tkinter.Tk()
root.withdraw()

# Creates excel workbook and sheet
workbook = xlsxwriter.Workbook('PATInformation.xlsx')
worksheet = workbook.add_worksheet()
# Adds headers for the columns
worksheet.write('A1','Instance')
worksheet.write('B1','Type')
worksheet.write('C1','Date')
# Initialize list that will hold instance,type,and date
list = []

# Displays file dialog
directory = tkFileDialog.askdirectory(initialdir="~/")

# If something was chosen...
if directory:
    # Go through each file in the directory that was chosen
    for filename in os.listdir(directory):
        path = directory + "/" + filename
        # Only parse files that end in .eml or .msg
        if (pathlib2.Path(path).suffix == ".eml" or pathlib2.Path(path).suffix == ".msg"):
            file = open(path)
            theEmail = email.message_from_file(file)
            file.close()

            # Gets the instance and date
            subject = theEmail['subject']
            instance = subject.split("-",1)[1].strip()
            date = theEmail['date']

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
        else:
            print ("The file " + filename + " is not a .eml or .msg so skipping")
row = 1
col = 0
for x in list:
    # Instance
    worksheet.write(row,col,x[0])
    # Type
    worksheet.write(row,col + 1,x[1])
    # Date
    worksheet.write(row,col + 2,x[2])
    row = row + 1
tkMessageBox.showinfo("PAT-Automation","Excel sheet has been created successfully!")
workbook.close()
