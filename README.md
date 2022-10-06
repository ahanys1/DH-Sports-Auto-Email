# DH-Sports-Auto-Email
## This script is now depreciated as it is no longer required.

Note: I am working on making this an application that will require no code editing to run. This is how it is for now.
## Setup
This is only done once.
### **Excel File Setup**

1. In the Email list First remove the heading row and ID number column. 

2. make sure that there are no blanks, and remove all rows that have blanks in them. 
3. Save the file. Once saved, DO NOT MOVE THE FILE LOCATION. 

### **Python Setup**

1. go to https://www.python.org/downloads/ and download version 3.10.1. Make sure to install pip with it. you do not need the IDLE. Notepad will work fine for what we need to do, however downloading Visual Studio Code (https://code.visualstudio.com/) might make it easier to run the program. 

2. Open Command Prompt and type the following Command:
    - pip install openpyxl

### **Code First Time Setup**

1. On line 10 replace "Insert path to email list here" with the file path to the email list. Use \\\ (double backslash) between folders (i.e. users\\\documents). Keep in quotes. 

2. On line 25 enter your email address and password in quotes.

3. On line 94 enter your email address.

### **Gmail Setup**

1. On Google, click your User Icon and select "Manage your Google Account".

2. Select "Security".

3. Turn on "Less Secure App Access".

## Setting Up the Script

This is done each time you run it. 
1. Download the ETA report workbook and open it.

2. In file explorer, find the location of the file. 

3. At the bar at the top, right click and select "Copy Address As Text". This will copy somthing like this to your clipboard: D:\abhmi\Documents

4. Paste this in line 11, keeping the quotes.

5. Add a second backslash after each one, and add another 2 at the end.

6. Type **the exact name** of the excel workbook, with .xlsx at the end of it, after what was pasted in. 

## Running the Script

### If using VSC
1. Click "Terminal" at the top, then "new terminal"

2. Make sure you are using Command prompt and not Powershell.

3. type "python emailer.py". This will run the program. DO NOT CLOSE ANYTHING UNTIL YOU SEE THE TIME ELAPSED. It will look like it's frozen for a while, but it is just loading the workbook.

### If not using VSC

1. Open Command prompt.

2. use the "cd" and "dir" commands to navigate to the folder containing "emailer.py".

3. type "python emailer.py". This will run the program. DO NOT CLOSE ANYTHING UNTIL YOU SEE THE TIME ELAPSED. It will look like it's frozen for a while, but it is just loading the workbook.
