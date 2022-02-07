# Resume-Filter-System

This project has the potential to be utilized as a resume management system. It can keep track of resumes, candidates' email addresses, and so on.

It takes the company's email address and password as input and searches for all unseen emails with topic keywords specified as parameters, such as employment or CV. It can also accept date or search before parameters.) It then requests weighted keywords to search for in the résumé (skillset). These skill sets are weighted and utilized to assign a score to each candidate.

This program may be executed as a simple Python script 'main.py.'
In the server parameter, enter your server's imap address. Go to gmail->my account->sign in and security->connected apps and site before launching the software. Allow less secure apps to be enabled.

It will download attachments from emails with specific topic keywords in a defined folder. The attachment might be in.pdf or.docx format. To avoid overwriting common names, resumes will be saved with the file name as the email id.

The resumes will next be scanned for the required expertise.(It may be tweaked for different applications.) and then scored.
How to examine a resume differs depending on the business.

The system will enter an email address, name, phone number, and date into an excel sheet and assign a 'yes' or 'no' decision to them.In the decision column, the computer asks for the keywords to search for in the resume and scores the applicant based on their weights and compares to the average of the, and if all of the keywords are present, the decision is yes; otherwise, the result is no.Candidates will receive an automatic response based on the choice in an excel document.

Every time the program is run, a new directory entitled current date-time is created in the directory resumes-and-candidate-data, and all resumes and the excel sheet are placed in that directory. This is done because every time a resume is downloaded via email, its status is changed from seen to unseen, and all of these resumes are placed in a specific date-time folder, avoiding duplication.

You must manually construct the'resumes-and-candidate-data' directory in the directory where the software is stored.
Here are several examples:

1. Every time the program is ran ,a new folder is created in the format shown below and each folder will have corresponding unseen resumes and the excel sheet with the candidate data and decisions.

![Alt text](https://github.com/Ojooh/Resume-Filter/blob/main/python%20auto%20recruitment%20pics/1.png'1')

2. When the program is run it initially asks for the input and then this kind of display is shown.The location of the the corresponding folder is shown, The resumes, number of words in resumes(just for debugging), and the enails to which replies are sent.

![Alt text](https://github.com/Ojooh/Resume-Filter/blob/main/python%20auto%20recruitment%20pics/2.png '2')

3. The next photo shows the contents of folder created above.

![Alt text](https://github.com/Ojooh/Resume-Filter/blob/main/python%20auto%20recruitment%20pics/3.png '3')

4. This is a sample of the excel sheet that the system will create.

![Alt text](https://github.com/Ojooh/Resume-Filter/blob/main/python%20auto%20recruitment%20pics/4.png '4')

5. This is an example of autmomated reply. The name of candidate will be taken from thte excel sheet.

![Alt text](https://github.com/Ojooh/Resume-Filter/blob/main/python%20auto%20recruitment%20pics/5.png '5')
