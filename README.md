# Gmail_Listener_and_Responder
As a Part of My Automated Account Ledger Downloading, Processing, and Delivery System for Dental Health Care Providers 

This tool is for use on Windows 10 systems which use Open Dental as their choice of practice management software

Running this Program:
- Will constantly monitor a user specified folder in your gmail inbox, awaiting emails upon refresh that have the subject format of "Lname Fname" and contian the word "ledger"
     - This was developed for automated, accurate ledger generation by front desk employees who had to bill bustomers who had a balance. This allowed them to receive a printable .xlsx file containing their entire formatted and formulated account ledger in under 2 minutes by sending an email to my inbox.
     
- Once one or a batch of emails containing this criteria are found, the information is sorted and dispatched to my Open Dental ledger creation export functions which produce a .xls file

- These are later passed to excel macros for formatting and formulating of the spreadsheet, and conversion of .xls to .xlsx

- This created ledger is then gathered by my send email function and sent to the email who sent the request.

# To Use:
Unless you run OD, much of this will not function correctly. Though the email scan and send functions are easily scalable to other applications of python :)

1. Download EmailMonitorTool.py
2. Enter your own email information and corresponding password in the file
3. Open cmd or terminal and cd to your downloads folder
      
      >>> python EmailMonitorTool.py
        
 # Example Output 
 (LOGIN FAILED!!! is generated every instance of refresh when there are no emails in the specified gmail inbox folder. Actual login fail will exit program)
 
![alt text](https://linkpicture.com/q/Screen-Shot-2020-06-25-at-12.14.32-PM.png)

