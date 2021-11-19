import pandas as pd
from google.colab import auth
import gspread
import string
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from oauth2client.client import GoogleCredentials
import time
import requests
import datetime
import hashlib
from google.colab import drive
auth.authenticate_user()
gc = gspread.authorize(GoogleCredentials.get_application_default())

#Checks if an email is valid or invalid using the "Real Email API"
def emailValid(email):
  api_key = API_KEY
  email_address = email
  response = requests.get(
      "https://isitarealemail.com/api/email/validate",
      params = {'email': email_address},
      headers = {'Authorization': "Bearer " + api_key })

  status = response.json()['status']
  if status == "invalid":
    return False
  else:
    return True

#Open Email Server
import smtplib
from email.mime.text import MIMEText
try:
    server = smtplib.SMTP_SSL(MAIL_HOST, MAIL_PORT)
    server.ehlo(MAIL_HOST)
    server.login(MAIL_USERNAME, MAIL_PASSWORD)
except:
    print ('Email server error...')

#Open Student Data Sheet and make Student_dataframe
try:
  Student_data_sheet = gc.open('Student Data Sheet').sheet1
except gspread.SpreadsheetNotFound:
  Print(month + ' Student Data spreadsheet not found')
Student_data = Student_data_sheet.get_all_values()
Student_dataframe = pd.DataFrame.from_records(Student_data, columns=['ID','Email','FirstName', 'LastName','Badge','ParentEmail','ParentApproval','TeacherEmail','TeacherApproval','ReviewerEmail','ReviewerApproval', 'LastRejected','LastReviewed','BadgeIssued'])

#Process a badge response form by emailing parents. This also processes Teachers because we need to know the badge the teacher approved so that we can send info to reviewers.
def processBadge(badge):
  global Student_dataframe
  try:
    Badge_data_sheet = gc.open(badge+' (Responses)').sheet1
  except gspread.SpreadsheetNotFound:
    print('Badge Data ' +badge+ ' spreadsheet not found')
    return
  Badge_data = Badge_data_sheet.get_all_values()

  # Since the number of questions a badge has varies I need to dynamically get the number of questions and append each one (Q1, Q2, etc) to the columnNames array
  secondHalf = Badge_data[0][8:]
  columnNames = ['Time','Email', 'FirstName','LastName','ParentEmail','TeacherEmail', 'Describe', 'Code']
  for i in range(1, len(secondHalf) + 1):
    qNum = "Q" + str(i)
    columnNames.append(qNum)

  Badge_dataframe = pd.DataFrame.from_records(Badge_data, columns = columnNames)

  # Removes duplicates and keeps only the latest submissions
  Badge_dataframe = Badge_dataframe.drop_duplicates(subset='Email', keep='last', ignore_index=True)

  #For each NEW badge application in a badge sheet, make a new entry in the Student Data Sheet, and email parent for confirmation
  for k in range(1,len(Badge_dataframe)):  #for all applications
    ID = hashlib.md5((str(Badge_dataframe['Email'][k])+badge).encode()).hexdigest()   #Student ID is hash of student email address + badge
    if ID not in Student_dataframe['ID'].values:   #k is a new application 
      parentValid = emailValid(Badge_dataframe['ParentEmail'][k]) #Validate parent email
      teacherValid = emailValid(Badge_dataframe['TeacherEmail'][k]) #Validate teacher email
      #Email the student if the parent or/and teacher's email is invalid. Also remove their row from the badge data sheet to allow a resubmission with the same ID
      if not parentValid or not teacherValid:
        email_intro = "<p>Hello. Please resubmit your application for the " + badge + " badge. The following emails are invalid:</p>"
        if not parentValid:
          email_intro = email_intro + "<h2>Parent's Email</h2> <p>" + Badge_dataframe['ParentEmail'][k] + "</p>"
        if not teacherValid:
          email_intro = email_intro + "<h2>Teacher's Email</h2> <p>" + Badge_dataframe['TeacherEmail'][k] + "</p>"
        email_intro = email_intro + "<h3>Here are your previous responses to use in the resubmission: </h3>"
        email_end = "<p>Please do not reply to this email. If you have questions, please use the contact information on the <a href='http://cs4ri.org/badges'>Rhode Island Computer Science Digital Badging website</a></p>"
        student_responses = "<ol>"
        for q in range(1,len(Badge_dataframe.columns)):  
          student_responses = student_responses + "<li><b>" + Badge_dataframe.iloc[0,q] + "</b><ul>" + Badge_dataframe.iloc[k,q] + "</ul></li> "
        email_body = email_intro + student_responses + "</ol><br>" + email_end
        error_email = MIMEText(email_body, "html")
        error_email["From"] = MAIL_FROM
        error_email["To"] = Badge_dataframe['Email'][k]
        error_email["Subject"] = badge + " submission error"
        server.sendmail(MAIL_FROM, Badge_dataframe['Email'][k], error_email.as_string())
        continue
    
      #make a new entry in student data frame and email parent
      Student_dataframe = Student_dataframe.append({'ID': ID,'Email': Badge_dataframe['Email'][k],'FirstName':Badge_dataframe['FirstName'][k],'LastName':Badge_dataframe['LastName'][k],'Badge':badge,'ParentEmail':Badge_dataframe['ParentEmail'][k],'ParentApproval':time.ctime(),'TeacherEmail':Badge_dataframe['TeacherEmail'][k],'TeacherApproval':'~','ReviewerEmail':'~','ReviewerApproval':'~','LastRejected':'~','LastReviewed':'~','BadgeIssued':'~'},ignore_index=True)
      email_body = "Hello. <br>" + Badge_dataframe['FirstName'][k] + " " + Badge_dataframe['LastName'][k] + " has applied for a <a href='http://cs4ri.org/badges'>Rhode Island Computer Science Digital Badge</a>. We ask that you fill out <a href='https://docs.google.com/forms/d/e/1FAIpQLSc_eDOkR5WCmffPLGQ7K2QrUet359Btk5nCibeGxD9SO6Y_xA/viewform'> this form </a> that asks you to allow us to review of this application. Please use the following in the form (we suggest pasting from this email into the form)<ul><b>Parent Email:</b> " + Badge_dataframe['ParentEmail'][k]+"<br><b>Student Key: </b>" + ID + "</ul>The review will be conducted by a trained member of our review team. No personal information, including the students name and email address will be shared with the reviewer. Only this automated Digital Badge system will communicate with you and your student to help them through the process of achieving their badge. Please do not reply to this email. If you have questions, please use the contact information on the <a href='http://cs4ri.org/badges'>Rhode Island Computer Science Digital Badging website</a>"
      parent_email = MIMEText(email_body, "html")
      parent_email["From"] = MAIL_FROM
      parent_email["To"] = Badge_dataframe['ParentEmail'][k]
      parent_email["Subject"] = "Computer Science Digital Badge Application Approval Needed For " + "  " + Badge_dataframe['FirstName'][k] + " " + Badge_dataframe['LastName'][k]
      server.sendmail(MAIL_FROM, Badge_dataframe['ParentEmail'][k], parent_email.as_string())
    else: # If reviewer rejected student application, send an email to student, reset reviewerapproval to default, and mark lastrejected as the current submission's timestamp to wait to send reviewer email until there's a new submission
      Student_sheet_index = Student_dataframe[Student_dataframe['ID']==ID].index.values[0]
      if Student_dataframe['ReviewerApproval'][Student_sheet_index] == False: 
        Student_dataframe['LastRejected'][Student_sheet_index] = Badge_dataframe['Time'][k]
        Student_dataframe['ReviewerApproval'][Student_sheet_index] = '~'

        try:
          Reviewer_data_sheet = gc.open('Reviewer Form (Responses)').sheet1
        except gspread.SpreadsheetNotFound:
          Print(month + ' Reviewer Response sheet not found')
        Reviewer_data = Reviewer_data_sheet.get_all_values()
        Reviewer_dataframe = pd.DataFrame.from_records(Reviewer_data, columns=['Time','Reviewer', 'Student','Approve','Comments'])  

        Reviewer_dataframe = Reviewer_dataframe.drop_duplicates(subset='Student', keep='last', ignore_index=True) # Removes duplicates to get the latest submission from the reviewer for that particular student

        email_intro = "<p>Hello. Your application for the " + badge + " badge has been denied. Please resubmit.</p>"
        email_intro = email_intro + "<h3>Here are your previous responses to use in the resubmission, plus the comments left by the reviewer:</h3>"
        email_end = "<p>Please do not reply to this email. If you have questions, please use the contact information on the <a href='http://cs4ri.org/badges'>Rhode Island Computer Science Digital Badging website</a></p>"
        student_responses = "<ol>"
        for q in range(1,len(Badge_dataframe.columns)):  
          student_responses = student_responses + "<li><b>" + Badge_dataframe.iloc[0,q] + "</b><ul>" + Badge_dataframe.iloc[k,q] + "</ul></li> "
        Reviewer_sheet_index = Reviewer_dataframe[Reviewer_dataframe['Student'] == ID].index.values[0]
        reviewer_comments = "<h2>Comments:</h2><p>" + Reviewer_dataframe['Comments'][Reviewer_sheet_index] + "</p>"
        email_body = email_intro + student_responses + "</ol><br>" + reviewer_comments + email_end
        error_email = MIMEText(email_body, "html")
        error_email["From"] = MAIL_FROM
        error_email["To"] = Badge_dataframe['Email'][k]
        error_email["Subject"] = badge + " denied"
        server.sendmail(MAIL_FROM, Badge_dataframe['Email'][k], error_email.as_string())

  #Email Reviewers for this badge with teacher approval that have not been emailed and have a valid email address in the Student Record sheet
  teacher_approved_indicies = Student_dataframe[(Student_dataframe['TeacherApproval']==True) & (Student_dataframe['Badge']==badge) & (Student_dataframe['ReviewerApproval']=='~') & (Student_dataframe['ReviewerEmail']!='~')].index.values  #list/array of all indicies in Student record with teacher approval for this badge where we need to email reviewer
  for i in teacher_approved_indicies:  
    Badge_sheet_index = Badge_dataframe[Badge_dataframe['Email'] == Student_dataframe['Email'][i]].index.values[0]
    if Badge_dataframe['Time'][Badge_sheet_index] != Student_dataframe['LastRejected'][i]: #Checks if this is a new submission (not rejected)
      Student_dataframe['ReviewerApproval'][i] = time.ctime()  #put into student record the date/time the reviewer was emailed
      email_intro = "Hello. <br> A student has applied for the <a href='http://cs4ri.org/badges'> "+badge+ " Rhode Island Computer Science Digital Badge</a>. We ask that you fill out <a href='https://docs.google.com/forms/d/e/1FAIpQLScFV2vLBMPBNpTeIWh2GVlwkpepW25Dm_KOWi-9TAKrZwDtFg/viewform'> this form </a>, which includes a rubric. Please use the following in the form (we suggest pasting from this email into the form)<ul><b>Reviewer Email Address:</b> " + Student_dataframe['ReviewerEmail'][i]+"<br><b>Student Key: </b>" + Student_dataframe['ID'][i]+ "</ul>Here are the student responses to assess in the form:<br>"
      email_end = "Please do not reply to this email. If you have questions, please use the contact information on the <a href='http://cs4ri.org/badges'>Rhode Island Computer Science Digital Badging website</a>"
      #Create an HTML formatted string of all questions in the form and all student responses
      #find the student's row in the badge response form
      student_responses = "<ol>"
      for q in range(6,len(Badge_dataframe.columns)):  #Assume all forms have questions starting in column 6 
        student_responses = student_responses + "<li><b>" + Badge_dataframe.iloc[0,q] + "</b><ul>" + Badge_dataframe.iloc[Badge_sheet_index,q] + "</ul></li> "
      email_body = email_intro + student_responses + "</ol><br>" + email_end
      reviewer_email = MIMEText(email_body, "html")
      reviewer_email["From"] = MAIL_FROM
      reviewer_email["To"] = Student_dataframe['ReviewerEmail'][i]
      reviewer_email["Subject"] = "Computer Science Digital Badge Application Review Needed For Student Key " + "  " + Student_dataframe['ID'][i]
      server.sendmail(MAIL_FROM, Student_dataframe['ReviewerEmail'][i], reviewer_email.as_string())
  
  #Look for reviewer approvals to write to CSV file for this badge
  reviewer_approved_indicies = Student_dataframe[(Student_dataframe['ReviewerApproval']==True) & (Student_dataframe['Badge']==badge)].index.values  #list/array of all indicies in Student record with reviewer approval for this badge where we need to email reviewer
  
  if len(reviewer_approved_indicies) > 0:   #Write CSV file
    drive.mount('/content/gdrive')
    with open('/content/gdrive/My Drive/' + badge + '.csv', 'w') as csvFile:
      csvFile.write("Email,FirstName,LastName,Narrative,NarrativeEvidence,EvidenceURL,IssueDate,ExpirationDate\n") 
      for k in reviewer_approved_indicies: #All who are approved - write CSV file
        #Email,	First Name,	Last Name, Narrative,	Narrative	Evidence = "Achieved proficiency according to the Rhode Island Computer Science Education standards", https://cs4ri.org/badges,	Issue Date,	Expiration Date
        csvFile.write(Student_dataframe['Email'][k] + "," + Student_dataframe['FirstName'][k] + "," + Student_dataframe['LastName'][k] + "," + "" + "," + "Achieved proficiency according to the Rhode Island Computer Science Education standards," + "https://cs4ri.org/badges ," + time.ctime() + "," + "" + "\n") 
    csvFile.close()
    Student_dataframe['BadgeIssued'][k] = time.ctime()  #mark time that badge was issued
  return

def processTeacher():
  global Student_dataframe
  try:
    global Teacher_data_sheet
    Teacher_data_sheet = gc.open('Teacher Form (Responses)').sheet1
  except gspread.SpreadsheetNotFound:
    Print(month + ' Teacher Response sheet not found')
    return
  Teacher_data = Teacher_data_sheet.get_all_values()
  global Teacher_dataframe
  Teacher_dataframe = pd.DataFrame.from_records(Teacher_data, columns=['Time','Teacher', 'Student','Confirm'])

  #Check for teacher approvals
  for j in range (1,len(Teacher_dataframe)):
    if (Teacher_dataframe['Student'][j] == '~'): #if submission was invalid (and has been "crossed out with ~") skip it
      continue
    if (len(Student_dataframe[Student_dataframe['ID']==Teacher_dataframe['Student'][j]].index.values) > 0): #if ID exists (typed in correctly)
      Student_sheet_index = Student_dataframe[Student_dataframe['ID']==Teacher_dataframe['Student'][j]].index.values[0]  #find index of student parent consented to in Student_dataframe
      if (Student_dataframe['TeacherApproval'][Student_sheet_index] == False): #if parent disapproved then skip student
        continue
      if (Teacher_dataframe['Teacher'][j] == Student_dataframe['TeacherEmail'][Student_sheet_index]): #verify that this is from the teacher
        if Teacher_dataframe['Confirm'][j] == 'Yes':
          Student_dataframe['TeacherApproval'][Student_sheet_index] = True  #Mark in student record that teacher approved
        elif Teacher_dataframe['Confirm'][j] == 'No':
          email_body = "<p>Hello. Confirmation for the " + Student_dataframe['Badge'][Student_sheet_index] + "badge has been denied by your teacher.</p>"
          student_email = MIMEText(email_body, "html")
          student_email["From"] = MAIL_FROM
          student_email["To"] = Student_dataframe['Email'][Student_sheet_index]
          student_email["Subject"] = Student_dataframe['Badge'][Student_sheet_index] + " badge confirmation denied"
          server.sendmail(MAIL_FROM, Student_dataframe['Email'][Student_sheet_index], student_email.as_string())
          Student_dataframe['TeacherApproval'][Student_sheet_index] = False  #Mark in student record that teacher disapproved
      else: #if email is invalid
        email_body = "<p>Hello. There was an error reading the teacher email from the Approval form. Please resubmit <a href='https://docs.google.com/forms/d/e/1FAIpQLScskZHSfoPOE6tXZJQ4qAL5pdKkcOjuqYkNz9toKsTiqe4Y3g/viewform'>here.</a> <strong>Carefully copy and paste the teacher email from the previous email.</strong></p>"
        teacher_email = MIMEText(email_body, "html")
        teacher_email["From"] = MAIL_FROM
        teacher_email["To"] = Student_dataframe['TeacherEmail'][Student_sheet_index]
        teacher_email["Subject"] = "Approval form submission error"
        server.sendmail(MAIL_FROM, Student_dataframe['TeacherEmail'][Student_sheet_index], teacher_email.as_string())
        for i in range(0,len(Teacher_dataframe.columns)):  #Fill invalid row with ~
          Teacher_dataframe.iloc[j,i] = '~'
    else: #if ID is invalid
      email_body = "<p>Hello. There was an error reading the student ID from the Approval form. Please resubmit <a href='https://docs.google.com/forms/d/e/1FAIpQLScskZHSfoPOE6tXZJQ4qAL5pdKkcOjuqYkNz9toKsTiqe4Y3g/viewform'>here.</a> <strong>Carefully copy and paste the student ID from the previous email.</strong></p>"
      teacher_email = MIMEText(email_body, "html")
      teacher_email["From"] = MAIL_FROM
      teacher_email["To"] = Teacher_dataframe['Teacher'][j]
      teacher_email["Subject"] = "Approval form submission error"
      server.sendmail(MAIL_FROM, Teacher_dataframe['Teacher'][j], teacher_email.as_string())
      for i in range(0,len(Teacher_dataframe.columns)):  #Fill invalid row with ~
        Teacher_dataframe.iloc[j,i] = '~'
  return

#Process parent responses. If approved, send to teacher. If rejected, destroy student data. 
def processParent():
  global Student_dataframe
  try:
    global Parent_data_sheet
    Parent_data_sheet = gc.open('Parent Form (Responses)').sheet1
  except gspread.SpreadsheetNotFound:
    Print(month + ' Parent Response sheet not found')
  Parent_data = Parent_data_sheet.get_all_values()
  global Parent_dataframe
  Parent_dataframe = pd.DataFrame.from_records(Parent_data, columns=['Time','Parent', 'Student','Permission'])
  
  for j in range(1,len(Parent_dataframe)):  #for all parent responses
    if (Parent_dataframe['Student'][j] == '~'): #if submission was invalid (and has been "crossed out with ~") skip it
      continue
    if (len(Student_dataframe[Student_dataframe['ID']==Parent_dataframe['Student'][j]].index.values) > 0): #if ID exists (typed in correctly)
      Student_sheet_index = Student_dataframe[Student_dataframe['ID']==Parent_dataframe['Student'][j]].index.values[0]  #find index of student parent consented to in Student_dataframe
      if (Student_dataframe['ParentApproval'][Student_sheet_index] == False): #if parent disapproved then skip student
        continue
      if (Parent_dataframe['Parent'][j] == Student_dataframe['ParentEmail'][Student_sheet_index]): #verify that this is from the parent
        # Also check if parent approval false in student sheet
        if Parent_dataframe['Permission'][j] == 'Yes':
          Student_dataframe['ParentApproval'][Student_sheet_index] = True  #Mark in student record that parent approved
          if Student_dataframe['TeacherApproval'][Student_sheet_index] == '~':  #Email teacher for approval
            Student_dataframe['TeacherApproval'][Student_sheet_index] = time.ctime()  #put into student record the date/time the teacher was emailed
            email_body = "Hello. <br>" + Student_dataframe['FirstName'][Student_sheet_index] + " " + Student_dataframe['LastName'][Student_sheet_index] + " has applied for the <a href='http://cs4ri.org/badges'> "+Student_dataframe['Badge'][Student_sheet_index]+ " Rhode Island Computer Science Digital Badge</a>. We ask that you fill out <a href='https://docs.google.com/forms/d/e/1FAIpQLScskZHSfoPOE6tXZJQ4qAL5pdKkcOjuqYkNz9toKsTiqe4Y3g/viewform'> this form </a> to verify that the student did the work with you. Please use the following in the form (we suggest pasting from this email into the form)<ul><b>Supervisor Email:</b> " + Student_dataframe['TeacherEmail'][Student_sheet_index]+"<br><b>Student Key: </b>" + Student_dataframe['ID'][Student_sheet_index]+ "</ul>Please do not reply to this email. If you have questions, please use the contact information on the <a href='http://cs4ri.org/badges'>Rhode Island Computer Science Digital Badging website</a>"
            teacher_email = MIMEText(email_body, "html")
            teacher_email["From"] = MAIL_FROM
            teacher_email["To"] = Student_dataframe['TeacherEmail'][Student_sheet_index]
            teacher_email["Subject"] = "Computer Science Digital Badge Application Approval Needed For " + "  " + Student_dataframe['FirstName'][Student_sheet_index] + " " + Student_dataframe['LastName'][Student_sheet_index]
            server.sendmail(MAIL_FROM, Student_dataframe['TeacherEmail'][Student_sheet_index], teacher_email.as_string())
        elif Parent_dataframe['Permission'][j] == 'No':
          email_body = "<p>Hello. Your permission to apply for the " + Student_dataframe['Badge'][Student_sheet_index] + "badge has been denied by your parent/guardian.</p>"
          student_email = MIMEText(email_body, "html")
          student_email["From"] = MAIL_FROM
          student_email["To"] = Student_dataframe['Email'][Student_sheet_index]
          student_email["Subject"] = Student_dataframe['Badge'][Student_sheet_index] + " badge permission denied"
          server.sendmail(MAIL_FROM, Student_dataframe['Email'][Student_sheet_index], student_email.as_string())
          Student_dataframe['ParentApproval'][Student_sheet_index] = False  #Mark in student record that parent disapproved
      else: #if email is invalid
        email_body = "<p>Hello. There was an error reading the parent email from the Approval form. Please resubmit <a href='https://docs.google.com/forms/d/e/1FAIpQLSc_eDOkR5WCmffPLGQ7K2QrUet359Btk5nCibeGxD9SO6Y_xA/viewform'>here.</a> <strong>Carefully copy and paste the parent email from the previous email.</strong></p>"
        parent_email = MIMEText(email_body, "html")
        parent_email["From"] = MAIL_FROM
        parent_email["To"] = Student_dataframe['ParentEmail'][Student_sheet_index]
        parent_email["Subject"] = "Approval form submission error"
        server.sendmail(MAIL_FROM, Student_dataframe['ParentEmail'][Student_sheet_index], parent_email.as_string())
        for i in range(0,len(Parent_dataframe.columns)):  #Fill invalid row with ~
          Parent_dataframe.iloc[j,i] = '~'
    else: #if ID is invalid
      email_body = "<p>Hello. There was an error reading the student ID from the Approval form. Please resubmit <a href='https://docs.google.com/forms/d/e/1FAIpQLSc_eDOkR5WCmffPLGQ7K2QrUet359Btk5nCibeGxD9SO6Y_xA/viewform'>here.</a> <strong>Carefully copy and paste the student ID from the previous email.</strong></p>"
      parent_email = MIMEText(email_body, "html")
      parent_email["From"] = MAIL_FROM
      parent_email["To"] = Parent_dataframe['Parent'][j]
      parent_email["Subject"] = "Approval form submission error"
      server.sendmail(MAIL_FROM, Parent_dataframe['Parent'][j], parent_email.as_string())
      for i in range(0,len(Parent_dataframe.columns)):  #Fill invalid row with ~
        Parent_dataframe.iloc[j,i] = '~'
  return


#Process Reviewer responses. If a student is accepted, write out Badgr CSV file.
def processReviewer():
  global Student_dataframe
  try:
    global Reviewer_data_sheet
    Reviewer_data_sheet = gc.open('Reviewer Form (Responses)').sheet1
  except gspread.SpreadsheetNotFound:
    Print(month + ' Reviewer Response sheet not found')
  Reviewer_data = Reviewer_data_sheet.get_all_values()
  global Reviewer_dataframe
  Reviewer_dataframe = pd.DataFrame.from_records(Reviewer_data, columns=['Time','Reviewer', 'Student','Approve','Comments'])  
  
  for r in range(1,len(Reviewer_dataframe)):  #for all reviewer responses
    if (Reviewer_dataframe['Student'][r] == '~'): #if submission was invalid (and has been "crossed out with ~") skip it
      continue
    if (len(Student_dataframe[Student_dataframe['ID']==Reviewer_dataframe['Student'][r]].index.values) > 0): #if ID exists (typed in correctly)
      Student_sheet_index = Student_dataframe[Student_dataframe['ID']==Reviewer_dataframe['Student'][r]].index.values[0]  # find index of student in Student_dataframe
      if (Reviewer_dataframe['Reviewer'][r] == Student_dataframe['ReviewerEmail'][Student_sheet_index]): #verify that this is from the Reviewer
        if Reviewer_dataframe['Approve'][r] == 'Yes' and Reviewer_dataframe['Time'][r] != Student_dataframe['LastReviewed'][Student_sheet_index]:
          Student_dataframe['ReviewerApproval'][Student_sheet_index] = True
          Student_dataframe['LastReviewed'][Student_sheet_index] = Reviewer_dataframe['Time'][r] #inserts time that the review was submitted into student dataframe to prevent any of the same changes when running the script again
          email_intro = "<p>Hello. Your application for the " + badge + " badge has been approved. You will receive an email from Badgr within the next week about your issued badge.</p>"
          email_end = "<p>Please do not reply to this email. If you have questions, please use the contact information on the <a href='http://cs4ri.org/badges'>Rhode Island Computer Science Digital Badging website</a></p>"
          reviewer_comments = "<h2>Comments:</h2><p>" + Reviewer_dataframe['Comments'][r] + "</p>"
          email_body = email_intro + reviewer_comments + email_end
          error_email = MIMEText(email_body, "html")
          error_email["From"] = MAIL_FROM
          error_email["To"] = Student_dataframe['Email'][Student_sheet_index]
          error_email["Subject"] = badge + " approved"
          server.sendmail(MAIL_FROM, Student_dataframe['Email'][Student_sheet_index], error_email.as_string())
        elif Reviewer_dataframe['Approve'][r] == 'No' and Reviewer_dataframe['Time'][r] != Student_dataframe['LastReviewed'][Student_sheet_index]:
          if (Reviewer_dataframe['Reviewer'][r] == Student_dataframe['ReviewerEmail'][Student_sheet_index]): #verify that this is from the Reviewer
            Student_dataframe['ReviewerApproval'][Student_sheet_index] = False
            Student_dataframe['LastReviewed'][Student_sheet_index] = Reviewer_dataframe['Time'][r] #inserts time that the review was submitted into student dataframe to prevent any of the same changes when running the script again
      else: #if email is invalid
        email_body = "<p>Hello. There was an error reading the reviewer email from the Approval form. Please resubmit <a href='https://docs.google.com/forms/d/e/1FAIpQLScFV2vLBMPBNpTeIWh2GVlwkpepW25Dm_KOWi-9TAKrZwDtFg/viewform'>here.</a> <strong>Carefully copy and paste the reviewer email from the previous email.</strong></p>"
        reviewer_email = MIMEText(email_body, "html")
        reviewer_email["From"] = MAIL_FROM
        reviewer_email["To"] = Student_dataframe['ReviewerEmail'][Student_sheet_index]
        reviewer_email["Subject"] = "Approval form submission error"
        server.sendmail(MAIL_FROM, Student_dataframe['ReviewerEmail'][Student_sheet_index], reviewer_email.as_string())
        for i in range(0,len(Reviewer_dataframe.columns)):  #Fill invalid row with ~
          Reviewer_dataframe.iloc[r,i] = '~'
    else: #if ID is wrong
      email_body = "<p>Hello. There was an error reading the student ID from the Approval form. Please resubmit <a href='https://docs.google.com/forms/d/e/1FAIpQLScFV2vLBMPBNpTeIWh2GVlwkpepW25Dm_KOWi-9TAKrZwDtFg/viewform'>here.</a> <strong>Carefully copy and paste the student ID from the previous email.</strong></p>"
      reviewer_email = MIMEText(email_body, "html")
      reviewer_email["From"] = MAIL_FROM
      reviewer_email["To"] = Reviewer_dataframe['Reviewer'][r]
      reviewer_email["Subject"] = "Approval form submission error"
      server.sendmail(MAIL_FROM, Reviewer_dataframe['Reviewer'][r], reviewer_email.as_string())
      for i in range(0,len(Reviewer_dataframe.columns)):  #Fill invalid row with ~
        Reviewer_dataframe.iloc[r,i] = '~'
  return


#### Main Program #####
processParent()
processTeacher()  # Do this before badges, because it sets teacher approval, which could trigger a badge being sent to a reviewer in badges
processReviewer() # Do this before badges, because it sets reviewer approval, which could trigger a writing of CSV in bages

# This array holds the names of each badge that you want to run through the script. The format is "<badge name> level <level #>".
badges = []
for badge in badges:
  processBadge(badge)


#Write out updated Student, Parent, Teacher and Reviewer data sheets from internal dataframes
set_with_dataframe(Student_data_sheet, Student_dataframe, include_column_header=False)
set_with_dataframe(Parent_data_sheet, Parent_dataframe, include_column_header=False)
set_with_dataframe(Teacher_data_sheet, Teacher_dataframe, include_column_header=False)
set_with_dataframe(Reviewer_data_sheet, Reviewer_dataframe, include_column_header=False)

server.close()
