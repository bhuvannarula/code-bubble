## examCollectionHelper v1.0

* demo files are present for your help. make sure to delete/modify them to your need before running program

### Origin of idea:
Due to COVID 19, my school had decided that subjective papers have to be mailed to class teacher in form of a .pdf file with student's name in the file name.
So i created this program to reduce the work of teachers.

### What it does?
Say 'Physics' exam was held today and student have sent their submissions.

- This program asks for the exam subject, exam date, and class.

- Creates a directory for that exam under 'student/' directory with format 'subject name-date-class' (eg: 'student/Physics-01-08-20-XII-A')

- Checks in 'studentDetails/' directory for a .csv file with filename as the class name containing the names of students of that class. (say 'XII-A.csv')

- Reads the credentials for email account of the class teacher given in 'cred.txt' and logs in to the email account using smtp to read the emails.

- Reads the emails received on the exam date and searches for attachments. 

- In case the filename of an attachment contains name of student present in the class list (in .csv file like - 'XII-A.csv'), it creates a folder under exam folder with having name the name of student and then downloads the attachment in this folder.

- In case two mails have been received with attachments from a student, the latter attachment is kept and older is over-written.

- When all mails have been searched, it asks if you want to export the data.

- if said no, you can re-run the program and it will check emails from the last email checked for more emails and keeps on getting data till the user does not exports. (say you downloaded data of 20 students out of 35, and the remaining sent the data later, you can re-run the program and it will download data of rest 15 and save it, without re-downloading the data of 20 students already downloaded)

- If said yes, it creates a studentList.csv file for subject teacher to write marks in (having cols 'Name' and 'Marks'), and writes '(not submitted)' in 'Marks' column for students whose email/attachment was not received, and then creates a zip file of this folder (eg: zip file is created of directory 'student/Physics-01-08-20-XII-A') under the directory 'archive/'

- Then asks if the files to be emailed to subject teacher.

- If said yes, it asks for the email of subject teacher, logs in to the email account using imap, and sends email to the subject teacher with subject as 'exam name-date-class' and the zip file as attachment.

### Sample Directory of data downloaded:

- archive/
	- Physics-01-08-20-XII-A.zip
	- Chemistry-01-08-20-XII-A.zip

- cred.txt

- student/
	- Physics-01-08-20-XII-A/
		- studentList.csv
		- Student Name 1
			- (Student Name 1's submission here)
		- Student Name 2
			- (Student Name 2's submission here)
	- Chemistry-01-08-20-XII-A/
		- studentList.csv
		- Student Name 1
			- (Student Name 1's submission here)
		- Student Name 2
			- (Student Name 2's submission here)

- studentDetails/
	- XII-A.csv

### Running this program:

*** delete following directories before running program:
'archive/'
'student/'

- create a csv file with student names of your class under 'studentDetails' folder.
- replace username and password in cred.txt file with that of your email account.
- for most emails, you need to change your email settings to allow access to unsecure applications (you can google it for how to do it).
- in this program it has been assumed that your email account is of gmail. if not so, replace the imap server name and smtp server name in main.py file with that of your email account (you can find it by google)
- run the main.py file and enter data as asked.
- internet connection is required for this program to run.
