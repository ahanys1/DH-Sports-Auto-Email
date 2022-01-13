import openpyxl

#file Locs
emailList = "D:\abhmi\Documents\test-email-book.xlsx"

#open workbooks
wbEmails = openpyxl.load_workbook(emailList)
wbEmailsSheet = wbEmails.active

#define area of emails list
emailsCells = wbEmailsSheet['A1':'C2']

#printing values of sheet
for row in emailsCells:
    for cell in row:
        print(cell.value)