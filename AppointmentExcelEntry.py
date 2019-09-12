#! python3
#Creator: Nathaniel Wiley
#Description: Updates excel spreadsheet with future tutoring appointments
#Runs every hour on the hours using windows task scheduler
import imaplib,re,datetime,sys
from openpyxl import Workbook,load_workbook
from tutFunctions import *
from privateInfo import *

#excel spreadsheet where appointment data is stored by this program
wb=load_workbook('C:\\Users\\Nate Wiley\\Documents\\TutoringAppointments.xlsx')
wks=wb.active

#checking if computer host is connected to internet
try:
    mail=imaplib.IMAP4_SSL('imap.gmail.com')
except:
    print('No internet connection')
    sys.exit()

try:
    F=wks['F1'].value #gets last run time as a datetime object from excel spreadsheet
    lastRunTime=F.strftime('%d')+'-'+F.strftime('%b')+'-'+str(F.year)
    mail.login(returnEmail(),returnEmailPassword())
    mail.select('inbox',readonly=True)
    SPECS='(SINCE "01-Jan-2019")'
    result, data = mail.search(None, SPECS)#gets emails since last run time
    if len(data)==0:#no new emails
        print('exited')
        sys.exit()
    email_id_list = data[0].split()
    for j in range(len(email_id_list)):
        a=getAppointmentList(email_id_list,j,mail)
        excelDate(a[0],a[1],wks)
    wks['F1']=datetime.datetime.today()#puts last run time in F1
    wb.save('C:\\Users\\Nate Wiley\\Documents\\TutoringAppointments.xlsx')
    mail.logout()
except:#lets user know what time an error ocurrs 
    print('Error')
    today=datetime.datetime.today()
    wks['F1']=today#puts last run time in F1 in excel
    wks['D1']='ERROR  '+today.strftime('%m %d %H:%M')
    wb.save('C:\\Users\\Nate Wiley\\Documents\\TutoringAppointments.xlsx')
    mail.logout()
