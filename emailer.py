from pickle import FALSE
import openpyxl
import smtplib
from email.message import EmailMessage


#file Locs
emailList = "D:\\abhmi\\Documents\\test-email-book.xlsx"
stockList = "D:\\abhmi\\Documents\\DH Open Sales Orders with ETAs January 11 2022 rerun TEST COPY.xlsx"

#open workbooks
wbEmails = openpyxl.load_workbook(emailList)
print("email list loaded")
wbEmailsSheet = wbEmails.active
wbStock = openpyxl.load_workbook(stockList)
print("stock list loaded")

#define area of emails list
emailsCells = wbEmailsSheet['A1':'B1']

#printing values of email sheet
for cell1, cell2, in emailsCells:
    print(cell1.value, cell2.value)
SelectedEmail = wbEmailsSheet["A1"].value

#select stock sheet from wbemails
wbStockSheet = wbStock.get_sheet_by_name(SelectedEmail)

#define area of stock list
stockCells = wbStockSheet['A1':'F250']

#printing values of stock sheet
message = ""
skip = False
for cell1, cell2, cell3, cell4, cell5, cell6 in stockCells:
    if cell1.value == None and cell2.value == None and cell3.value == None and cell4.value == None and cell5.value == None and cell6.value == None and skip == True:
        break
    elif cell1.value == None and cell2.value == None and cell3.value == None and cell4.value == None and cell5.value == None and cell6.value == None and skip == False:
        skip = True
    
    if (cell1.value != None):
        one = cell1.value
        one = str(one)
        if one[0] in "0123456789":
            one = (one[:10])
        if isinstance(one, str):
            one = one.ljust(19,".")
    else:
        one = "..................."
    if (cell2.value != None):
        two = cell2.value
        two = two.ljust(7, ".")
    else:
        two = "......."
    if (cell3.value != None):
        three = cell3.value
        three = three.ljust(11,".")
    else:
        three = "..........."
    if (cell4.value != None):
        four = cell4.value
        four = four.ljust(46, ".")
    else:
        four = "."*46
    if (cell5.value != None):
        five = cell5.value
        five = str(five)
        five = five.ljust(14,".")
    else:
        five = "................"
    if (cell6.value != None):
        six = cell6.value
        six = str(six)
    else:
        six = "     "
    #print(one, two, three, four, five, six)
    message = message + one +" "+ two +" "+ three +" "+ four +" "+ five +" "+ six +" "+ "\n" 
    print (message)
message = "<pre>" + message + "</pre>"
msg = EmailMessage()
msg.set_content(message)

msg['Subject'] = 'please work'
msg['From'] = "abhmind@gmail.com"
msg['To'] = "abhmind@gmail.com"

# Send the message via our own SMTP server.
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login("abhmind@gmail.com", "fierland")
server.send_message(msg)
server.quit()