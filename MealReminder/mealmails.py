#!python2

import pdb
import datetime
import csv
import sys
from os import path

from includes.objects import Student, Group, MailServer

# constants as defaults
FROM = "REU Meal Shift System"
SUBJECT = "Meal Shift Reminder"

# address of server to connect to
SMTP_SERVER = 'smtp.fas.harvard.edu'

# directory where data files are stored
DATA_DIR = path.join(path.dirname(path.realpath(sys.argv[0])),"data")

def loadStudentInfo(filePath,group):
  '''
  Reads filePath, which should be a csv with two columns, student name and student
  email. Creates student objects for each student and adds them to the REU group
  so that they are searchable.
  '''
  with open(filePath) as studentFile:
    # skip first line
    next(studentFile)

    # create student object per line and add it to group
    if all(map(lambda info: group.addMember(Student(info[0], info[1])),
           csv.reader(studentFile))):
      return group
    else:
      return None

def prompt(default,message):
  '''
  Promts the users with message (no additional characters). If user enters nothing,
  default is returned, otherwise the user input is returns.
  '''
  inp = input(message)
  return (inp if inp != '' else default)

def main():
  '''
  Runs the main program to automatically read in files and set out the information.
  '''
  # defaults
  dFrom, dmailFile, dtimeFile = "REUMeals@meals.harvard.edu", "REU Student Emails.csv","Meal Shifts 2014.xlsx"
  dHours = 12
  dtimeName = "Signs"
  demailTemp = "template.txt"
  hemailTemp = "template.html"

  # messages
  mFrom = "From email? ({}): ".format(dFrom)
  mmailFile = "Email Address File ({}): ".format(dmailFile)
  mtimeFile =  "Time Sheet File ({}): ".format(dtimeFile)
  mtimeName = "Time Sheet Name ({}): ".format(dtimeName)
  mHours = "Warn within how many hours? ({}): ".format(dHours)
  memailTemp = "Text Email template? ({})".format(demailTemp)
  hmemailTemp = "HTML Email template? ({})".format(hemailTemp)

  # prompt for modifications
  (From, mailFile, timeFile, tHours, timeSheet, emailTemp, htmlTemp) = (prompt(dFrom,mFrom),
                                                              prompt(dmailFile,mmailFile),
                                                              prompt(dtimeFile,mtimeFile),
                                                              int(prompt(dHours, mHours)),
                                                              prompt(dtimeName, mtimeName),
                                                              prompt(demailTemp, memailTemp),
                                                              prompt(hemailTemp, hmemailTemp))

  # construct paths to files
  mailPath, emailPath, htmlPath, timePath = (path.join(DATA_DIR, mailFile),
                                                  path.join(DATA_DIR, emailTemp),
                                                  path.join(DATA_DIR, htmlTemp),
                                                  path.join(DATA_DIR, timeFile)) 
  #pdb.set_trace()
  # read emails into classes
  REUGroup = loadStudentInfo(mailPath, Group(eFile=timePath, eSheet=timeSheet))
  #pdb.set_trace()
  # read student times within the written range
  dutySet = REUGroup.upcomingMembersDuties(time=datetime.datetime.now(), 
                                           timeRange=datetime.timedelta(hours=tHours))

  #open connection to mailServer
  server = MailServer(SMTP_SERVER)
  try:
    # create email and send out for each person
    templates = {}
    templates['text'] = open(emailPath).read().replace("[from]", FROM).replace("[hours]",str(tHours))
    templates['html'] = open(htmlPath).read().replace("[from]", FROM).replace("[hours]",str(tHours))

    # pdb.set_trace()
    for (name, duties) in dutySet.items():
      # get required information and check for existense of member
      if REUGroup.isMember(name):
        fullName = REUGroup.fullName(name)
        email = REUGroup.members[fullName].email
        content = server.createMessageContent(name,duties, templates)
        #pdb.set_trace()
        # send it out
        server.sendemail(content, email, From, SUBJECT)

        # tell the user
        print("Sent email to {}({}) for {} upcoming shift(s).".format(name, email, len(duties)))
      else:
        print("Error. Attempted to send email to {}, but they are not in the group.".format(
              name))
        print(duties)

  finally:
    # close connection
    server.quit()

if __name__ == '__main__':
  main()

  # wait for user input
  input("\n\nPress Enter to exit!\n\n")
