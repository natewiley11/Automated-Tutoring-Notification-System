#! python3
#Creator: Nathaniel Wiley
#Sends text message with appointment info via Twilio 
#Runs every half hour on the half hours using windows task scheduler
from twilio.rest import Client
from openpyxl import Workbook,load_workbook
import datetime,sys
from privateInfo import *

#twilio
#These functions refer to a private file that I did not include in this project
#for privacy purposes
accountSID=returnAccountSID()
authToken=returnAuthToken
twilioClnt=Client(accountSID,authToken)
myTwilioNum=returnMyTwilioNum()
myPhoneNum=returnPhoneNum()

#Checking excel file for imminent (within 40 min) appointments
wb=load_workbook('C:\\Users\\Nate Wiley\\Documents\\TutoringAppointments.xlsx')
wks=wb.active
AppDate=wks['B2'].value

if AppDate==None:
    print('No appointments')
    sys.exit()
Subj=wks['C2'].value
if AppDate.hour<12:
    dateString=AppDate.strftime('%D %H:%M')
else:
    x=AppDate.hour-12
    dateString=AppDate.strftime('%D '+str(x)+':%M')
msg='Appointment\n'+dateString+'\nSubject: '+Subj
if AppDate-datetime.timedelta(minutes=40)<datetime.datetime.today():
    #sends the message if appointment is within 40 min of now
    wks.delete_rows(2,1)#deletes appointment
    message=twilioClnt.messages.create(body=msg,
                                       from_=myTwilioNum,to=myPhoneNum)


wb.save('C:\\Users\\Nate Wiley\\Documents\\TutoringAppointments.xlsx')
