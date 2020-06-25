import sys
import imaplib
import getpass
import email
import email.header
import datetime
import time
import os
from os import listdir
from os.path import isfile, join
import pyautogui
import time
import openpyxl
from PIL import ImageGrab
from PIL import Image
import win32com.client as win32
import glob, shutil
import csv
import smtplib
import base64
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


EMAIL_ACCOUNT = "stevens@######"
EMAIL_PASS = '########'
EMAIL_FOLDER = 'Auto' #selects auto-filtered inbox. Configure separately in gMail.IF NO EMAIL IN BOX, OUTPUTS LOGIN ERROR
M = imaplib.IMAP4_SSL('imap.gmail.com')

def process_mailbox(M):
    email_bank = []

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return

        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        sender = email.header.make_header(email.header.decode_header(msg['From']))
        #date = email.header.make_header(email.header.decode_header(msg['Date']))
        #print(str(date))
        #print(sender)
        subject = str(hdr)
        email_bank.append(str(num)+': '+subject+' from '+str(sender))
        M.store(num, '+X-GM-LABELS', '\\Trash')
        #print('%s: %s' % (num, subject))
    return(email_bank)

def get_g(string):
    if string.find('-') == -1:
        return('F')
    i = string.find('-')
    i = i + 2
    gname = string[i:]
    end_i = gname.find(' from')
    g_name = gname[:end_i]
    return(g_name)

def get_s(string):
    if string.find('<') == -1:
        return('F')
    i = string.find('<')
    i = i + 1
    ename = string[i:]
    end_i = ename.find('>')
    e_name = ename[:end_i]
    return(e_name)

def get_emails():
    global EMAIL_ACCOUNT
    global EMAIL_PASS
    global EMAIL_FOLDER
    global M

    export = []
    try:
        rv, data = M.login(EMAIL_ACCOUNT, EMAIL_PASS)
    except imaplib.IMAP4.error:
        print ("LOGIN FAILED!!! ")
        exit()

    rv, mailboxes = M.list()
    rv, data = M.select(EMAIL_FOLDER)

    if rv == 'OK':
        email_list = process_mailbox(M)
        for x in email_list:
            #get guarantor name to lookup
            guarantor = get_g(x)
            ##print("guarantor name list output below:")
            #print(guarantor)
            if guarantor == 'F':
                #print("INVALID ENTRY FORMAT...Email Subject format is 'Audit Request - Fname Lname'")
                #put function here to email back with instructions on input
                ########
                ########
                continue
            #get email address of sender for respective request
            sender = get_s(x)
            ##print("sender email below:")
            #print(sender)
            string = guarantor+'>'+sender
            export.append(string)
        #M.close()
        #M.logout()
        return(export);

    else:
        print("ERROR: Unable to open mailbox ", rv)
        #M.close()
        #M.logout()
        return('F')

def check_emails():
    data = get_emails()
    data_bank = []
    for d in data:
        #print(d)
        data_l = d.split('>')
        data_bank.append(data_l)
        #print(data_l)
    return(data_bank)

def parser(emails):
    #gets name
    name = emails[0].split(' ')
    #print(str(name))
    fname = name[0]
    lname = name[1]
    emailz = emails[1]

    return (fname, lname, emailz)

#This will use OpenDental, and open it if it is not, login with provided user credentials, and export account aging ledgers
#It sources the list of patients to monitor from Resources\CV_tracker.xlsx
#Will run through list once. If you want automation, run it from a scheduled batch file or shell script.
#The program will work with any operating system, but requires correct images for the repective system to know where to locate buttons.

export_path = 'C:\\Users\\ASD_Staff\\Desktop\\Resources\\Exported'
root_path = 'C:\\Users\\ASD_Staff\\Desktop\\Resources'
data_path = 'C:\\Users\\ASD_Staff\\Desktop\\Resources\\Data'
existing_path = 'C:\\Users\\ASD_Staff\\Dropbox\\Collections, EP\\Small_Claim\\Steven\\Pre Judgment'

def startup():

    odup = -1

    while odup == -1:
        try:
            report_button0 = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Reports_button_off.png'))
            odup = 0
            report_x = report_button0[0] + 20
            report_y = report_button0[1] + 10
        except TypeError:
            break

    while odup == -1:
        try:
            report_button1 = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Reports_button.png'))
            odup = 1
            report_x = report_button1[0] + 20
            report_y = report_button1[1] + 10
        except TypeError:
            break

    while odup == -1:
        try:
            pyautogui.keyDown('win'); pyautogui.press('r'); pyautogui.keyUp('win')
            time.sleep(1)
            pyautogui.typewrite('C:\\Program Files (x86)\\Open Dental\\OpenDental.exe')
            pyautogui.press('enter')
            odup = 2
        except TypeError:
            break

    #login if just opened from code
    if odup == 2:
        time.sleep(5)
        login_select = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\User_s.png'))
        user_x = login_select[0] + 10
        user_y = login_select[1] + 5
        pyautogui.click(user_x, user_y, clicks=2, interval=1)
        pyautogui.typewrite('4899654')
        pyautogui.press('enter')
        time.sleep(5)
        odup = 1

    if odup == -1:
        print('Did not assign odup value properly during initial sorting')
        exit()
    if odup == 2:
        print('Could not find report button on screen')
        exit()
#add check for trojan update and waiting period here
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    while odup == 0 or odup == 1:
        #once lonched w need to find report button again, used previously as odup reference (instead of writing another verification method)
        while odup == 0:
            try:
                report_button0 = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Reports_button_off.png'))
                odup = 9
                report_x = report_button0[0] + 20
                report_y = report_button0[1] + 10
                break
            except TypeError:
                break

        while odup == 1:
            try:
                report_button1 = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Reports_button.png'))
                odup = 9
                report_x = report_button1[0] + 20
                report_y = report_button1[1] + 10
                break
            except TypeError:
                break
    return(odup, report_x, report_y);
    #ODUP return tells whether OD window opened sucessfully

#clicks report button when OD is up

def report_button(report_x, report_y):
    pyautogui.click(report_x, report_y, clicks=2, interval=0.25)
    time.sleep(2)
    q_button = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\UserQ_button.png'))
    query_x = q_button[0] + 0
    query_y = q_button[1] + 5
    pyautogui.click(query_x, query_y, clicks=2, interval=0.25)
    q_selection = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Query_select_button.png'))
    query_x = q_selection[0] + 30
    query_y = q_selection[1] + 8
    pyautogui.click(query_x, query_y, clicks=2, interval=0.25)
    return();


def download(lname, fname):
    try:
        status = list(startup())
        status_od = status[0]
        report_x = status[1]
        report_y = status[2]
        if status_od != 9:
            print('ODUP not set to 9')
            return(odup);

        #OpenDental is up and running now
        time.sleep(1)

        try:
            guarantor = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\select_pt.png'))
            g_x = guarantor[0] + 12
            g_y = guarantor[1] + 12
            error = 0
        except TypeError:
            error = 1
        if error == 1:
            guarantor = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\select_pt_on.png'))
            g_x = guarantor[0] + 12
            g_y = guarantor[1] + 12
        pyautogui.click(g_x, g_y, clicks=1)
        time.sleep(1)
        pyautogui.typewrite(lname)
        time.sleep(0.25)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.typewrite(fname)
        time.sleep(0.5)
        archive = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\archive_on.png'))
        arch_x = archive[0] + 5
        arch_y = archive[1] + 5
        pyautogui.click(arch_x, arch_y, clicks = 1)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)

    #checks for popup on account
        try:
            popup_detect = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\popup_detect.png'))
            get_off = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\popup_ok.png'))
            off_x = get_off[0] + 50
            off_y = get_off[1] + 20
            pyautogui.click(off_x, off_y, clicks = 2)
            time.sleep(1)
        except TypeError:
            print('')
        try:
            family = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\family.png'))
            f_x = family[0] + 12
            f_y = family[1] + 12
            errorr = 0
        except TypeError:
            errorr = 1
        try:
            pt_error_check = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\no_pt.png'))
            return(-9);
        except TypeError:
            print('')

        if errorr == 1:
            family_on = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\family_on.png'))
            f_x = family_on[0] + 12
            f_y = family_on[1] + 12
        pyautogui.click(f_x, f_y, clicks=1)
        time.sleep(3)
        pt_info = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\pt_info_entry.png'))
        f_x = pt_info[0] + 12
        f_y = pt_info[1] + 50
        pyautogui.click(f_x, f_y, clicks=2, interval=0.1)
        time.sleep(2)
        pnum_location = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\pnum_loc.png'))
        f_x = pnum_location[0] + 125
        f_y = pnum_location[1] + 5
        pyautogui.click(f_x, f_y, clicks=2, interval=0.2)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'c')
        cancel = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\cancel.png'))
        cancel_x = cancel[0] + 12
        cancel_y = cancel[1] + 25
        pyautogui.click(cancel_x, cancel_y, clicks=2, interval=0.2)
        time.sleep(1)

    #with debtor selected, run report query now
        report_button(report_x, report_y)

        time.sleep(1)
        guarantor = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\set_guarantor.png'))
        g_x = guarantor[0] + 225
        g_y = guarantor[1] + 7
        pyautogui.click(g_x, g_y, clicks=1)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        g_set = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\g_entry.png'))
        ok_x = g_set[0] + 40
        ok_y = g_set[1] + 20
        pyautogui.click(ok_x, ok_y, clicks = 1)
        time.sleep(5)
        export = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\export_b.png'))
        exp_x = export[0] + 20
        exp_y = export[1] + 8
        pyautogui.click(exp_x, exp_y, clicks = 1)
        time.sleep(2)

    #global use
        global export_path
        path = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\save_path.png'))
        path_x = path[0] + 400
        path_y = path[1] + 20
        pyautogui.click(path_x, path_y, clicks = 1)
        pyautogui.typewrite(export_path)
        time.sleep(0.2)
        pyautogui.press('enter')

        save_as = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\save_xl.png'))
        s_x = save_as[0] + 90
        s_y = save_as[1] + 11
        pyautogui.click(s_x, s_y, clicks = 1)
        time.sleep(1)
        pyautogui.press('down')
        pyautogui.press('enter')

        name = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\file_name.png'))
        name_x = name[0] + 120
        name_y = name[1] + 10
        pyautogui.click(name_x, name_y, clicks = 1)

            #saves as .xls for use in later processing
        name_str = "{} {} Ledger CV - Raw XLS".format(lname, fname)
        pyautogui.typewrite(name_str)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1)

        try:
            conf_save = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Confirm_save.png'))
            f_x = conf_save[0] + 50
            f_y = conf_save[1] + 20
            pyautogui.click(f_x, f_y, clicks = 2)
            pyautogui.press('left')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(1)
        except TypeError:
            print('')

        pyautogui.hotkey('alt', 'f4')
        pyautogui.hotkey('alt', 'f4')

        try:
            popup_detect = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Excel_repl_yes.png'))
            off_x = popup_detect[0] + 50
            off_y = popup_detect[1] + 20
            pyautogui.click(off_x, off_y, clicks = 2)
            pyautogui.press('left')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(1)
        except TypeError:
            print('')


    ###########################################
    ###########################################
    # DONE WITH OPENDENTAL, NOW DATA PROCESSING
        xl = win32.gencache.EnsureDispatch('Excel.Application')

    #save .xls as editable .xlsx
        source_file = "{}\\{}.xls".format(export_path, name_str)
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        xl.DisplayAlerts = False
        xl.Visible = False
        wb = xl.Workbooks.Open(source_file)
        print("Source File is:")
        print(source_file)

        wb.SaveAs(source_file+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
        wb.Close()                                                  #FileFormat = 56 is for .xls extension
        xl.Application.Quit()
        time.sleep(1)

        try:
            xlapp = win32.gencache.EnsureDispatch('Excel.Application')
        except AttributeError:
                # Corner case dependencies.
                import os
                import re
                import sys
                import shutil
                # Remove cache and try again.
                MODULE_LIST = [m.__name__ for m in sys.modules.values()]
                for module in MODULE_LIST:
                    if re.match(r'win32com\.gen_py\..+', module):
                        del sys.modules[module]
                shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        #from win32com import client
        #xlapp = win32.gencache.EnsureDispatch('Excel.Application')
        #apply macro to new .xlsx file and save it
        #xlapp = win32.gencache.EnsureDispatch('Excel.Application')
        xlapp.DisplayAlerts = False
        xlapp.Visible = False

        xlbook = xlapp.Workbooks.Open('C:\\Users\\ASD_Staff\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.xlsb')
        xlbook2 = xlapp.Workbooks.Open(source_file+"x")

        xlapp.Application.Run("PERSONAL.XLSB!DeleteFinance")
        time.sleep(1)
        xlapp.Application.Run("PERSONAL.XLSB!MakePrettyandFillFormulas")
        time.sleep(5)

        xlapp.ActiveWorkbook.SaveAs(Filename = source_file+"x")

        xlbook2.Save()
        xlbook.Close(SaveChanges=True)
        xlbook2.Close(SaveChanges=True)
        xlapp.Quit()

        del xlbook
        del xlbook2
        del xlapp


        try:
            popup_detect = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Excel_repl_yes.png'))
            off_x = popup_detect[0] + 50
            off_y = popup_detect[1] + 20
            pyautogui.click(off_x, off_y, clicks = 2)
            pyautogui.press('left')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(1)
        except TypeError:
            print('')

        try:
                conf_save = list(pyautogui.locateOnScreen('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Images\\Confirm_save.png'))
                f_x = conf_save[0] + 50
                f_y = conf_save[1] + 20
                pyautogui.click(f_x, f_y, clicks = 2)
                pyautogui.press('left')
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(1)
        except TypeError:
            print('')

        except IndexError:
            return(1);
    except:
        return(1);


def killxl():
    pyautogui.keyDown('win'); pyautogui.press('r'); pyautogui.keyUp('win')
    time.sleep(1)
    pyautogui.typewrite("cmd")
    pyautogui.keyDown('ctrl'); pyautogui.keyDown('shift'); pyautogui.press('enter'); pyautogui.keyUp('ctrl'); pyautogui.keyUp('shift')
    time.sleep(2)
    pyautogui.typewrite("taskkill /F /IM excel.exe"); pyautogui.press('enter')
    time.sleep(2)
    return(1)

def run_cleanup():

    v = killxl()
    if v != 1:
        print("Could not launch admin cli to kill Excel")
        exit()

def get_bal_sum(fdir, fname):
    file_path = fdir+"\\"+fname
    #file_open = open('C:\\Users\\ASD_Staff\\Desktop\\Resources\\Data\\Howe Josephine Ledger CV - Raw XLS_csv.csv')
    file_open = open(file_path)
    file_read = csv.reader(file_open)
    file_data = list(file_read)
    x = -1

    for row in file_data:
        x = x + 1

    while x > 10:
        string = str(file_data[x])
        if string.find(",,,,,,,,,,,") == -1:
            bal_sum = file_data[x]
            break
        x = x - 1

    return(bal_sum)

def get_name(file):
    name = ""
    j = 0
    x = 0
    for a in file:
        x = x + 1
    x = x - 1
    while file[x] != '\\':
        x = x - 1
    x = x + 1
    return(file[x:])

def is_odd(i):
    if i % 2 == 0:
        return False
    return True

def cleanup(lname, fname):

    global export_path

    file = export_path+"\\"+lname+' '+fname+' '+"Ledger CV - Raw XLS.xlsx"
    return(1);

def send_email(lname, fname, email):

    global export_path

    filename = export_path+"\\"+lname+' '+fname+' '+"Ledger CV - Raw XLS.xlsx"

    fromaddr = "stevens@#####.net"
    toaddr = email

    # instance of MIMEMultipart
    msg = MIMEMultipart()

    # storing the senders email address
    msg['From'] = fromaddr

    # storing the receivers email address
    msg['To'] = toaddr

    # storing the subject
    msg['Subject'] = "Your requested Account Ledger"

    # string to store the body of the mail
    body = "THE OD AUDINATOR      Copyright 2019 Steven Steinbeck"

    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent
    #filename = "File_name_with_extension"
    attachment = open(filename, "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
    p.set_payload((attachment).read())

    # encode into base64
    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login(fromaddr, 'Password_HERE')

    # Converts the Multipart msg into a string
    text = msg.as_string()

    # sending the mail
    s.sendmail(fromaddr, toaddr, text)

    # terminating the session
    s.quit()

def remove_existing(lname, fname):

    #XL file removal from Export Dir
    global data_path
    global raw_stream
    global export_path
    global root_path

    try:
        string = export_path+"\\"+lname+' '+fname+' '+"Ledger CV - Raw XLS.xlsx"
        os.remove(string)
        stringd = export_path+"\\"+lname+' '+fname+' '+"Ledger CV - Raw XLS.xls"
        os.remove(stringd)

    except FileNotFoundError:
        print("No old ledger found >>> Generating new one")

def send_error(lname, fname, email):

    fromaddr = "stevens@####.net"
    toaddr = email

    # instance of MIMEMultipart
    msg = MIMEMultipart()

    # storing the senders email address
    msg['From'] = fromaddr

    # storing the receivers email address
    msg['To'] = toaddr

    # storing the subject
    msg['Subject'] = "Unable to Process your requested Account Ledger"

    # string to store the body of the mail
    body = "THE OD AUDINATOR requires a valid patient name. {} {} is not valid. Please check the submitted credentials and try again.              Copyright 2019 Steven Steinbeck".format(fname, lname)

    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login(fromaddr, 'Password_HERE')

    # Converts the Multipart msg into a string
    text = msg.as_string()

    # sending the mail
    s.sendmail(fromaddr, toaddr, text)

    # terminating the session
    s.quit()

def run_specific_report(fname, lname, email):

    #runs download loop for all debtors in CV_tracker.xlsx file
    v = killxl()
    if v != 1:
        print("Excel could not be terminated. Goodbye.")
        sys.exit()
    remove_existing(lname, fname)
    status = download(lname, fname)
    time.sleep(10)
    if status == -9:
        send_error(lname, fname, email)
        print("Input Error notification sent sucessfully :)")
    else:
        status_3 = send_email(lname, fname, email)
        print("Data was downloaded, modified, and e-mailed sucessfully :)")



#MAIN
#MAIN
#MAIN
#constantly monitors email for new guarantor
while True:
    try:
        emails = check_emails()
        if emails == 0:
            print("No emails found, check Gmail folder")
            sys.exit()
        for e in emails:
            #print(e)
            if e[0].find('INVAL') != -1:
                emails.remove(e)
            fname, lname, emailz = parser(e)
            print("Generating Report For:")
            print(fname, lname)
            print("To be sent to:")
            print(emailz)
            run_specific_report(fname, lname, emailz)
            time.sleep(30)
        time.sleep(60)
        exit()
    except:
        time.sleep(60)
        continue

#MODIFY XL VBA SCRIPT TO VARIABLY SELECT RANGE FOR SORTING
#ADD SUPPORT FOR NON-THEMED WINDOWS GUI
#ADD SUPPORT FOR ESC OFFICE SEARCH
#  -requires additional support for combining XL sheets before sorting with VBA scripts.
#ADD CHECK FOR EXISTING HUMAN-MODIFIED LEDGER IN LEDGER DIRECTOY. SKIP DOWNLOAD AND PROCESSING, JUST EMAIL.
