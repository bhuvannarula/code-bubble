import tkinter
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
from time import sleep

'''
Help

Version 2.0

Change Log:
- added GUI to the program. Program is now User-Friendly
- bug fixes
- some improvements

* Don't panic if program freezes. It is working and doing it's task. Let it finish.

* MAKE SURE YOU ENABLE ACCESS TO UNSECURE APPS FROM YOUR EMAIL ACCOUNT
    (in case it shows any login error and the cred.txt is correctly filled.)

* demo files are present for your help. make sure to delete/modify them to your need before running program

* read README.md for how to run this program

*** delete following directories before running program:
'archive/'
'student/'

* if it shows wrong password error, remove any new lines in cred.txt

'''


win1 = tkinter.Tk('examCollectionHelper','examCollectionHelper','examCollectionHelper')
bgcolor = 'Black'
fgcolor = 'White'
borderwidth1 = 0

#declaring variables
className1=tkinter.StringVar()
examDate1=tkinter.StringVar()
examSubject1=tkinter.StringVar()
subjectTeacherMail=tkinter.StringVar()
checkZipCondition1 = tkinter.BooleanVar()
checkZipCondition1.set(True)

frame1 = tkinter.Frame(win1,bg=bgcolor,width=500,height=380)
frame1.pack()

labelHeading = tkinter.Label(frame1,text='examCollectionHelper',font=('Helventica',25,'bold'),fg=fgcolor,bg=bgcolor)
labelWarning1 = tkinter.Label(frame1,font=('Helventica',10),text='* Make sure you have updated Email ID, Password in cred.txt . Read README.md for any help. *',fg=fgcolor,bg=bgcolor)
labelHeading.place(x=10,y=10)
labelWarning1.place(x=10, y=50)

labelClassname = tkinter.Label(frame1,text='Class Name (eg: "XII-SA1"):',font=('Helventica',15),fg=fgcolor,bg=bgcolor)
labelClassname.place(x=10,y=95)
inputClassname = tkinter.Entry(textvariable=className1,width=15,borderwidth = borderwidth1)
inputClassname.place(x=220,y=95)

labelExamSubject = tkinter.Label(frame1,text='Exam Subject:',font=('Helventica',15),fg=fgcolor,bg=bgcolor)
labelExamSubject.place(x=10,y=135)
inputExamSubject = tkinter.Entry(textvariable=examSubject1,width=15,borderwidth = borderwidth1)
inputExamSubject.place(x=130,y=135)

labelExamDate = tkinter.Label(frame1,text='Exam Date (dd/mm/yy):',font=('Helventica',15),fg=fgcolor,bg=bgcolor)
labelExamDate.place(x=10,y=175)
inputExamDate = tkinter.Entry(textvariable=examDate1,width=15,borderwidth = borderwidth1)
inputExamDate.place(x=195,y=175)

inputSubjectTeacherMail = tkinter.Entry(frame1,borderwidth=borderwidth1,textvariable=subjectTeacherMail,width=25)
def toggleSTM():
    inputSubjectTeacherMail['state'] = 'normal' if checkZipCondition1.get() else 'disabled'

labelZipCondition = tkinter.Label(frame1,text='Do you want to export and mail data?',font=('Helventica',15),fg=fgcolor,bg=bgcolor)
checkZipCondition = tkinter.Checkbutton(frame1,command=toggleSTM,variable=checkZipCondition1,onvalue=True,offvalue=False,bg=bgcolor,fg=fgcolor)
labelZipCondition.place(x=10,y=245)
checkZipCondition.place(x=275,y=247)


labelSubjectTeacherMail = tkinter.Label(frame1,text='Subject Teacher Email ID:',font=('Helventica',15),fg=fgcolor,bg=bgcolor)
#inputSubjectTeacherMail = tkinter.Entry(frame1,textvariable=subjectTeacherMail,width=25,state='disabled)
#inputSubjectTeacherMail['state']='normal'
labelSubjectTeacherMail.place(x=10,y=285)
inputSubjectTeacherMail.place(x= 210, y=285)

from tkinter import messagebox

def fm1():
    print(messagebox.showerror('Error!','Wrong Email/Password'))



def raiseError(message1):
    messagebox.showerror('Error!',message1)

def askUser(message1):
    return messagebox.askyesno('Warning!',message1)


def mainloop():
    examSubject = examSubject1.get()
    className = className1.get()
    examDate = examDate1.get()
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
    # take 'Subject Name' and 'Class' and 'Exam Date'
    #examSubject = input('Enter Subject Name: ')
    #className = input('Enter Class (Demo: "XII-SA1"): ')
    #examDate = input('Enter Exam Date (like dd/mm/yy): ')

    if examSubject == '' or className == '' or not re.search(r'../../..$',examDate):
        raiseError('No Subject or Class given or Wrong date format.')
        return None
    examDate = examDate.replace('/','-')
    
    if not os.path.isfile('studentDetails/'+ className + '.csv'):
            raiseError('Student Details csv file does not exist for this class! Please go to folder: studentDetails/ and create .csv file for this class.')
            return None      
    
    examDateObject = datetime.strptime(examDate,'%d-%m-%y')
    examSubject2 = examSubject+' Exam-{}-{}'.format(examDate,className)

    # checks if the data for this class,exam has been sent/exported already
    if os.path.isdir(os.getcwd()+'/archive'):
        alreadyExported = os.listdir(os.getcwd()+'/archive')
        if (examSubject2+'.zip') in alreadyExported:
            resp1 = askUser('The file for the class,exam has already been exported. Do you still want to continue?')
            if not resp1:
                raiseError('the program has been stopped. restart if needed.')
                return None
            else:
                os.remove(os.getcwd()+'/archive/'+examSubject2+'.zip')
                shutil.rmtree(os.getcwd()+'/student/' + examSubject2)

    if not os.path.isdir(os.getcwd()+'/student/' + examSubject2):
        os.mkdir(os.getcwd()+'/student/' + examSubject2)

    # create folder for particular exam,date and class
    filePath1 = os.getcwd()+'/student/' + examSubject2 + '/'



    # this function creates a list of names of student from the studentDetails class csv file. 

    def prepareStudent2():
        try:
            fileIn = open('studentDetails/' +  className+'.csv','r')
        except FileNotFoundError:
            raiseError('Student Details csv file does not exist for this class! Please go to folder: studentDetails/ and create .csv file for this class.')
            return None
        csvFileIn = csv.reader(fileIn)
        studentDetailsdict = [item55[0] for item55 in list(csvFileIn)]
        return studentDetailsdict
    studentDetailsdict = prepareStudent2()
    if not studentDetailsdict:
        return None
    # these 3 lines read username, password from cred.txt
    credIn = open('cred.txt','r',).read().split(',')
    credIn = [item40.split('=')[1] for item40 in credIn]
    username,password = credIn
    try:
        # these 3 lines connect to mail account
        imap = imaplib.IMAP4_SSL("imap.gmail.com") # assuming email id is of gmail
        imap.login(username, password)
        status, messages = imap.select("INBOX")
    except:
        raiseError('Wrong Email Account Username/Password.')
        return None

    messages = int(messages[0])

    # function to check if student name is in class list
    def findStudentName(string):
        for item21 in studentDetailsdict:
            item22 = item21.lower().split(' ')
            if re.search('.*'.join(item22),string.lower()):
                return item21.title()

    # kind of deprived, was to be used when checking email id's to identify student
    def noAttachmentError(studentName):
        print("{}'s Email has no attachment!!".format(studentName))

    # following function decodes the mail message and gets the attachments
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
    #print(messages,cache[0])    

    # the main loop    
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

    # following function creates a csv file for subject teacher to write marks in
    def prepareStudent():
        try:
            fileIn = open('studentDetails/' +  className+'.csv','r')
        except FileNotFoundError:
            raiseError('Student Details csv file does not exist for this class!')
            return None
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
        return 1

    # this part exports the folder into zip file
    #exportZIPconditon = (input('Do you want to export zip file? (y/n): ').lower() in ('yes','y'))
    exportZIPconditon = checkZipCondition1.get()
    if exportZIPconditon:
        if not prepareStudent():
            return None
        if not os.path.isdir(os.getcwd()+'/archive'):
            os.mkdir(os.getcwd()+'/archive')
        output_filename = 'archive/' + examSubject2
        dir_name = os.getcwd() + '/student/' + examSubject2
        shutil.make_archive(output_filename, 'zip', dir_name)
        os.remove(os.getcwd() + '/cookies.bin')

    #if not (input('Do you want to email files? (y/n): ').lower() in ('yes','y'))
    if not exportZIPconditon:
        raiseError('All mails recieved till now have been checked. Re-run the program to check new mails.')
        return None

    # this inputs the email id of subject teacher
    teacherEmail = subjectTeacherMail.get()
    if not re.search('.+@.+\..*',teacherEmail):
        raiseError('Not a valid Subject Teacher Email ID')
        #teacherEmail = input("Enter {} Teacher Email ID: ".format(examSubject))
        return None


    # composing mail
    sendmsg = EmailMessage()
    sendmsg['Subject'] = examSubject2 + ' Files'
    sendmsg['From'] = username
    sendmsg['To'] = teacherEmail
    #zipattach = zipfile.ZipFile('archive/'+examSubject2+'.zip','r')
    zipattach = open('archive/'+examSubject2+'.zip','rb')
    sendmsg.add_attachment(zipattach.read(),maintype='application',subtype='zip',filename=(examSubject2+'.zip'))

    # following code sends the mail and logs out
    smtp = smtplib.SMTP_SSL('smtp.gmail.com')
    smtp.login(username,password)
    smtp.send_message(sendmsg)
    smtp.quit()
    del username
    del password
    
    messagebox.showinfo('Program Successful!','The files have been sent successfully. Have a nice day!')
    sleep(1.5)
    win1.destroy()

finalSubmitButton = tkinter.Button(frame1,command=mainloop,bg=bgcolor,text='Submit',borderwidth=borderwidth1)
finalSubmitButton.place(x=30,y=325)

tkinter.mainloop()