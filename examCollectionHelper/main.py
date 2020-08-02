import imaplib
import email
from email.message import EmailMessage
import os
import pickle
import re
import csv
from datetime import datetime
import shutil
import smtplib
import zipfile


'''
Help

* demo files are present for your help. make sure to delete/modify them to your need before running program

* read README.md for how to run this program

*** delete following directories before running program:
'archive/'
'student/'

* if it shows wrong password error, remove any new lines in cred.txt

'''

# open cache file which gives the last email read so as to prevent redoing
cacheFile = open('cookies.bin','ab+')
cacheFile.seek(0)
try:
    cache = pickle.load(cacheFile) # [last email viewed]
except:
    cache = [0]


# creates folder named 'student'
if not os.path.isdir(os.getcwd()+'/student/'):
    os.mkdir(os.getcwd()+'/student/')
# take 'Subject Name' and 'Class'
examSubject = input('Enter Subject Name: ')
className = input('Enter Class (Demo: "XII-SA1"): ')
examDate = input('Enter Exam Date (like dd/mm/yy): ')

if examSubject == '' or className == '' or not re.search(r'../../..$',examDate):
    raise UserWarning('No Subject or Class given or Wrong date format.')
examDate = examDate.replace('/','-')
examDateObject = datetime.strptime(examDate,'%d-%m-%y')
dateToday = datetime.strftime(datetime.now(),'%d-%m-%y')
dateToday =  datetime.strptime(dateToday,'%d-%m-%y')
examSubject2 = examSubject+' Exam-{}-{}'.format(examDate,className)

if os.path.isdir(os.getcwd()+'/archive'):
    alreadyExported = os.listdir(os.getcwd()+'/archive')
    if (examSubject2+'.zip') in alreadyExported:
        resp1 = input('The file for the class,exam has already been exported. Do you still want to continue? (y/n)').lower() in ('y','yes')
        if not resp1:
            raise UserWarning('the program has been stopped. restart if needed.')
        else:
            os.remove(os.getcwd()+'/archive/'+examSubject2+'.zip')
            shutil.rmtree(os.getcwd()+'/student/' + examSubject2)

if not os.path.isdir(os.getcwd()+'/student/' + examSubject2):
    os.mkdir(os.getcwd()+'/student/' + examSubject2)
# create folder for particular exam,date and class
filePath1 = os.getcwd()+'/student/' + examSubject2 + '/'



# this function makes a dictionary with key as student name and value as student's potential email(s)
# this also creates a csv file with student names in which teacher can enter marks 

def prepareStudent2():
    try:
        fileIn = open('studentDetails/' +  className+'.csv','r')
    except FileNotFoundError:
        raise UserWarning('Student Details csv file does not exist for this class!')
    csvFileIn = csv.reader(fileIn)
    studentDetailsdict = [item55[0] for item55 in list(csvFileIn)]
    return studentDetailsdict
studentDetailsdict = prepareStudent2()

'''
def prepareStudent():
    try:
        fileIn = open('studentDetails/' +  className+'.csv','r')
    except FileNotFoundError:
        raise UserWarning('Student Details csv file does not exist for this class!')
    fileOut2 = open('student/' + examSubject2 + '/studentList.csv','w',newline='')
    csvFileIn = csv.reader(fileIn)
    csvFileOut = csv.writer(fileOut2)
    csvFileOut.writerow(['Name of Student','Marks'])
    for item in csvFileIn:
        csvFileOut.writerow([item[0]])
    fileOut2.close()
studentDetailsdict = prepareStudent()
'''

# these 3 lines read username, password from cred.txt
credIn = open('cred.txt','r',).read().split(',')
credIn = [item40.split('=')[1] for item40 in credIn]
username,password = credIn

# these 3 lines connect to mail account
imap = imaplib.IMAP4_SSL("imap.gmail.com") # assuming email id is of gmail
imap.login(username, password)
status, messages = imap.select("INBOX")

messages = int(messages[0])
def findStudentName(string):
    for item21 in studentDetailsdict:
        item22 = item21.lower().split(' ')
        if re.search('.*'.join(item22),string.lower()):
            return item21.title()

def noAttachmentError(studentName):
    print("{}'s Email has no attachment!!".format(studentName))

def getAttachment(msg):
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            emailDate = datetime.strptime(re.search(r'([0-9]{2} .* [0-9]{4})',msg['date']).group(),'%d %b %Y')
            if emailDate > examDateObject:
                return 'new'
            elif emailDate < examDateObject:
                return 'old'
            #from_ = msg.get("From").split('<')[-1].split('>')[0]
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    try:
                        filename = part.get_filename()
                        if filename:
                            studentName = findStudentName(filename)
                            if not os.path.isdir(filePath1+studentName):
                                os.mkdir(filePath1+studentName)
                            filename = studentName + ' ' + className + '.' + filename.split('.')[-1]
                            filepath = os.path.join(filePath1+studentName, filename)
                            open(filepath, "wb").write(part.get_payload(decode=True))
                        else:
                            noAttachmentError(studentName)
                            return None
                    except:
                        None   
print(messages,cache[0])        
for i in range(messages, cache[0], -1):
    res, msg = imap.fetch(str(i), "(RFC822)")
    resp = getAttachment(msg)
    if resp == 'new':
        continue
    elif resp == 'old':
        break

imap.close()
imap.logout()

# Enter into cache how many emails have been read
cacheFile.seek(0)
cache[0] = messages
pickle.dump(cache,cacheFile)
cacheFile.flush()

def prepareStudent():
    try:
        fileIn = open('studentDetails/' +  className+'.csv','r')
    except FileNotFoundError:
        raise UserWarning('Student Details csv file does not exist for this class!')
    fileOut2 = open('student/' + examSubject2 + '/studentList.csv','w',newline='')
    csvFileIn = csv.reader(fileIn)
    csvFileOut = csv.writer(fileOut2)
    csvFileOut.writerow(['Name of Student','Marks'])
    for item in csvFileIn:
        if not os.path.isdir('student/'+examSubject2+'/'+item[0]):
            csvFileOut.writerow([item[0].title(),'(not submitted)'])
        else:
            csvFileOut.writerow([item[0].title()])
    fileOut2.close()

# this part exports the folder into zip file
exportZIPconditon = (input('Do you want to export zip file? (y/n): ').lower() in ('yes','y'))
if exportZIPconditon:
    prepareStudent()
    if not os.path.isdir(os.getcwd()+'/archive'):
        os.mkdir(os.getcwd()+'/archive')
    output_filename = 'archive/' + examSubject2
    dir_name = os.getcwd() + '/student/' + examSubject2
    shutil.make_archive(output_filename, 'zip', dir_name)
    os.remove(os.getcwd() + '/cookies.bin')

if not (input('Do you want to email files? (y/n): ').lower() in ('yes','y')):
    raise UserWarning('the program has been stopped. restart if needed.')

teacherEmail = input("Enter {} Teacher Email ID: ".format(examSubject))
while not re.search('.+@.+\..*',teacherEmail):
    print('Not a valid ')
    teacherEmail = input("Enter {} Teacher Email ID: ".format(examSubject))


# composing mail
sendmsg = EmailMessage()
sendmsg['Subject'] = examSubject2 + ' Files'
sendmsg['From'] = username
sendmsg['To'] = teacherEmail
#zipattach = zipfile.ZipFile('archive/'+examSubject2+'.zip','r')
zipattach = open('archive/'+examSubject2+'.zip','rb')
sendmsg.add_attachment(zipattach.read(),maintype='application',subtype='zip',filename=(examSubject2+'.zip'))

smtp = smtplib.SMTP_SSL('smtp.gmail.com')
smtp.login(username,password)
smtp.send_message(sendmsg)
smtp.quit()
del username
del password
