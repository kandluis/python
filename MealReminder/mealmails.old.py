#/usr/bin/env python

import smtplib
import re
import csv
import xlrd
import pdb

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime

from includes.objects import Student, Group, Duty, MailServer

# address of server to connect to
SMTP_SERVER = 'smtp.fas.harvard.edu'

# regex for email checks
EMAIL_REGEX = re.compile(r"[^@]+@[^@]+\.[^@]+")

# regex for dates of the form mm/dd
DATE_REGEX = re.compile(r"^(0?[1-9]|1[0-2])/(0?[1-9]|[12][0-9]|3[01])$")

# specifies tolerance for error when matching names
ERR_TOL = 0.50 # value from 1 to 2, with 1 requiring perfect match, and 2 not comparing

# contants for the date,row tuple we use
DATE = 0
ROW = 1

# constants as defaults
FROM = "REU Meal Shift System"
SUBJECT = "Meal Shift Reminder"
FROM = "Friends in F House"

class Duty(object):
  def __init__(self,rows,meal,mtype, timeString):
    self.rows = rows
    self.meal = meal
    self.type = mtype
    self.time = timeString

  def name(self):
    return(self.meal + " " + self.type)

EARLIEST_TIME = (7,0)
LATEST_TIME = (18,0)
# maps times to specific columns in the excel timesheet data as well as descriptors
TIMECOLS = {  "7:00AM"  : Duty([1], "Breakfast", "Prep.", "7:00 AM"),
              "8:00AM"  : Duty([2], "Breakfast", "Clean-Up", "8:00 AM"),
              "11:30AM" : Duty([3], "Lunch", "Prep.", "11:30 AM"),
              "12:30AM" : Duty([4,5], "Lunch", "Clean-Up", "12:30 AM"),
              "5:00PM"  : Duty([6], "Dinner", "Prep.", "5:00 PM"),
              "6:00PM"  : Duty([7,8], "Dinner", "Clean-Up", "6:00 PM")
}

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Helper Functions Used in Script - Should be Factored out if using in a larger
project.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
def checkTime(limBelow, limAbove):
  '''
  Returns the list of (index, duty) of the columns whose time is limBelow <= time <= limAbove
  @limAbove and limBelow are both (hh, mm) tuple.
  '''
  def underLim(t):
    '''
    Retunrs true if the time t (which is a string in format HH:MM{PM/AM}) occurs
    before lim.
    '''
    hhA, mmA = limAbove
    hhB, mmB = limBelow
    # get time for t and for lim
    tDate = datetime.datetime.strptime(t, "%I:%M%p")
    limAboveDate = datetime.datetime(year=tDate.year, month=tDate.month, day=tDate.day,
                                hour=hhA, minute=mmA)
    limBelowDate = datetime.datetime(year=tDate.year, month=tDate.month, day=tDate.day,
                                hour=hhB, minute=mmB)
    return(limBelowDate <= tDate <= limAboveDate)

  # list of duties
  lDuties = [duty for (time, duty) in TIMECOLS.items() if underLim(time)]
  return([(index, duty) for duty in lDuties for index in duty.rows])

def prompt(default,message):
  '''
  Promts the users with message (no additional characters). If user enters nothing,
  default is returned, otherwise the user input is returns.
  '''
  inp = input(message)
  return (inp if inp != '' else default)

def editDistance(s,t):
  '''
  Calculates the smallest number of deletions, insertions, or letter changes 
  required to convert the string s into the string t
  '''
  # zero index represents empty string, otherwise strings are 1-indexed
  results = [[max(i,j) for j in range(len(t)+1)] for i in range(len(s)+1)]

  # build bottom up using DP
  for i in range(1,len(s)+1):
    for j in range(1,len(t)+1):
      # in order of min - either delete character, insert character, match character, or change character
      match = 0 if s[i-1] == t[j-1] else 1
      results[i][j] = min(results[i-1][j] + 1, results[i][j-1] + 1, results[i-1][j-1] + match)

  return results[len(s)][len(t)]

def stringMatchDec(s1,s2):
  '''
  Returns how closely two strings match in decimal format
  '''
  return (editDistance(s1,s2)/max(len(s1),len(s2)))

def ithWord(s,i):
  '''
  Attempts to return i-th indexed word in a string. Indexes behave as they would
  if s was a list of words that is 1-indexed. Returns None if it can't index
  correctly. An empty string maps to the empty list of words.
  '''
  try:
    return(s.split(' ')[i-1] if len(s) != 0 else [])
  except IndexError:
    return None

def minIndex(lst):
  '''
  Returns the index of the minum value in the list
  '''
  return min(range(len(lst)), key=lst.__getitem__)

def cleanName(name): 
  '''
  Removes all non-characters from name
  '''
  # remove extraneous symbols and lowercase
  cName = re.sub(r'\W+( \W+)*$', '', name).upper()

  # some names have an extra space at the end
  try:
    if cName[-1] == " ":
      return(cName[0:len(cName)-1])
    return(cName)
  except IndexError:
    return(cName)

def numWords(s):
  return(len(s.split(" ")))

def loadStudentInfo(fileName,group):
  '''
  Reads fileName, which should be a csv with two columns, student name and student
  email. Creates student objects for each student and adds them to the REU group
  so that they are searchable.
  '''
  with open(fileName) as studentFile:
    # skip first line
    next(studentFile)

    # create student object per line and add it to group
    if all(map(lambda info: group.addMember(Student(info[0], info[1])),
           csv.reader(studentFile))):
      return group
    else:
      return None

class MailServer(object):
  '''
  A mail server is basically a connection to the server specified by URL which 
  allows us to send and create messages
  '''
  def __init__(self, url):
    self.server = smtplib.SMTP(url)

  def sendemail(self, message,To,From,Subject):
    '''
    Function send a message. A connection to @sever is assumed. 
    params:
    @message - the message to be sent, in string format
    @To - the email address the message should be sent to
    @From - the email address from which the message is sent@
    @Subject - the subject of the message
    '''
    msg = MIMEText(message)

    # set metadata details
    msg['Subject'] = Subject
    msg['From'] = From
    msg['To'] = To

    # send the message view SMTP server
    self.server.send_message(msg)

  def createMessage(self,name, duties, template):
    '''
    Returns the template with the appropriate name and duties inserted
    '''
    # construct string of duties
    dStr = ""
    for date,duty in duties:
      dStr += duty.time + " on " + date + ": " + duty.meal + " " + duty.type + "\n\n"

    return(template.replace('[shortname]', name).replace("[duties]", dStr))

  def quit(self):
    self.server.quit()

class Student(object):
  '''
  A student class to keep track of first name, last name, and email for a student
  '''
  def __init__(self,name,email):
    '''
    Student name and email in default input formats. 
    '''
    self.name = cleanName(name)
    self.email = email if EMAIL_REGEX.match(email) else None

class Group(object): 
  ''' 
  A group class keeps track of a group of students, and provides membership and 
  lookup functions for students 
  '''
  def __init__(self,name = "REU",eFile = "", eSheet = ""): 
    self.name = name 
    self.size = 0
  
    # keep track of student objects by full names 
    self.members = {} 
  
    # keep track of names of student objects 
    self.memberNames = []  

    # external data on members stored here
    self.extFile = eFile
    self.extSheet = eSheet
  
  def addMember(self,member): 
    ''' 
    Creates a link between fullname and the student object. Returns True if  
    succesfull, False if student already exists. Cleans the name.  
    '''
    cName = cleanName(member.name)
    if cName not in self.members and cName not in self.memberNames: 
      self.members[cleanName(cName)] = member 
      self.memberNames.append(cName)
      self.size += 1 
      return True
    else: 
      return False
  
  def isMember(self,name): 
    ''' 
    Returns true if name partially matches a memeber, false otherwise 
    '''
    return (self.fullName(name) is not None) 
  
  def fullName(self,name): 
    ''' 
    Attempts to partially match name to a member of the group. Partial matching is as follows: 
    1. Attempts full match with non-characters removed with error tolerance. 
    2. Attempts match from beginning of name to end with error tolerance. If at least entire first word 
        is matched, this is a partial match. 
    3. Attempts match from end of name to beginning with error tolerance. It at least entire last word 
        is matched, this is a partial match 
  
    Return the name of the partially matched meber, or returns None if no member matches 
    '''
    #pdb.set_trace()
    # clean name 
    cName = cleanName(name) 
      
    # returns list of matched items based on index 
    def matchList(i):
      '''
      Special list to match a specific word in a name with the passed in parameter.
      ''' 
      if i == -1: 
        return list(map(lambda t : stringMatchDec(cName,t), self.memberNames)) 
      else: 
        return list(map(lambda t : stringMatchDec(ithWord(cName,i), t),  
                        map(lambda s: ithWord(s,i), self.memberNames))) 
  
    # match fullnames if longer than one words
    if numWords(cName) > 1:
      flNameMatchDec = matchList(-1) 
      flNameInd = minIndex(flNameMatchDec)
      if flNameMatchDec[flNameInd] < ERR_TOL:  
        return self.memberNames[flNameInd]

    # single word name, so match first and last names  
    # match firstnames 
    fNameMatchDec = matchList(1)  
    fNameInd = minIndex(fNameMatchDec)

    # match lastnames
    lNameMatchDec = matchList(0) 
    lNameInd = minIndex(lNameMatchDec) 

    # return closest matching
    if fNameMatchDec[fNameInd] < lNameMatchDec[lNameInd] and fNameMatchDec[fNameInd] < ERR_TOL: 
      return self.memberNames[fNameInd]
    elif lNameMatchDec[lNameInd] < ERR_TOL: 
      return self.memberNames[lNameInd]
    else:
      return None

  def findMemberr(self,name): 
    ''' 
    Finds a member and returns the member object if found. Otherwise, returns None 
    '''
    if self.isMember(name): 
      return self.members[self.fullName(name)] 
    else: 
      return None

  def upcomingMembersDuties(self,time,timeRange):
    '''
    Returns a list of the upcoming student objects within the specified timerange.
    '''
    def inRange(dateStr):
      '''
      Input is of the form mm/dd. Return true only if it is in fact within the specified
      time range.
      '''
      # create artifical datetime for the date we are reading
      try:
        month, day = dateStr.split("/")
        E_hr, E_mm = EARLIEST_TIME
        date = datetime.datetime(year=time.year, month=int(month), day=int(day),
                                 hour=E_hr, minute=E_mm)
        return(time < date < (time + timeRange))
      
      # if string is not splittable or if dateStr is not a string with a split attribute
      except (ValueError,AttributeError) as e:
        return False

    def checkDates(sheet):
      '''
      Input a sheet where dates are in first column. Scan column for indexes of
      appropriate rows matchings the correct span of days. ithWord(x,0) brings you
      the 0th word 
      '''
      dataColumn = sheet.col_values(0)
      return([(i, ithWord(str(val),0)) for (i,val) in enumerate(dataColumn) if inRange(ithWord(str(val),0))])

    def readNamesDuties(rows):
      '''
      rows - the rows of the excel sheet that from which we need to grab names. The rows
      are
      The functions returns the text values of the cells inside rows that are within the 
      time range. It is assumed that rows contains each row, organized from earlier to later
      dates. The function returns a dictionary of name: (date,duty) list. Any empty 
      cells are ignored and note included in the dictionary.
      '''
      def accept(val):
        return (val != '')
      
      def add(d,k,v):
        if k in d:
          d[k].append(v)
        elif accept(k):
          d[k] = [v]

      # get limits
      lim = time + timeRange
      limBelow = (time.hour, time.minute)
      limAbove = (lim.hour, lim.minute)

      # create the dictionary to hold 
      matchings = {}

      if len(rows) == 1:
        for (i,duty) in checkTime(limBelow, limAbove):
          add(matchings,rows[0][ROW][i], (rows[0][DATE],duty))
      # multiple
      elif len(rows) > 1:
        # first row
        for (i,duty) in checkTime(limBelow, LATEST_TIME):
          add(matchings,rows[0][ROW][i], (rows[0][DATE], duty))

        # last row
        for (i,duty) in checkTime(EARLIEST_TIME,limAbove):
          add(matchings, rows[-1][ROW][i], (rows[-1][DATE], duty))

        # everything in between
        allDuties = checkTime(EARLIEST_TIME,LATEST_TIME)
        for rowI in range(1,len(rows)-1):    
          for (i,duty) in checkTime(EARLIEST_TIME,LATEST_TIME):
            add(matchings,rows[rowI][ROW][i], (rows[rowI][DATE], duty))
          
      return(matchings)

    # read correct sheet
    workbook = xlrd.open_workbook(self.extFile)
    signs = workbook.sheet_by_name(self.extSheet)

    # collect the row indexes and values to look at by analyzing first column
    rowIndVal = checkDates(signs)
    rowColl = [(val, signs.row_values(i)) for (i,val) in rowIndVal]

    # grab names and duty from the excel spredsheet, and return them
    return(readNamesDuties(rowColl))

def main():
  '''
  Runs the main program to automatically read in files and set out the information.
  '''
  # defaults
  dFrom, dmailFile, dtimeFile = "sknapp@fas.harvard.edu", "REU Student Emails.csv","Meal Shifts 2014.xlsx"
  dHours = 12
  dtimeName = "Signs"
  demailTemp = "template.txt"

  # messages
  mFrom = "From email? ({}): ".format(dFrom)
  mmailFile = "Email Address File ({}): ".format(dmailFile)
  mtimeFile =  "Time Sheet File ({}): ".format(dtimeFile)
  mtimeName = "Time Sheet Name ({}): ".format(dtimeName)
  mHours = "Warn within how many hours? ({}): ".format(dHours)
  memailTemp = "Email template? ({})".format(demailTemp)

  # prompt for modifications
  (From, mailFile, timeFile, tHours, timeSheet, emailTemp) = (prompt(dFrom,mFrom),
                                                              prompt(dmailFile,mmailFile),
                                                              prompt(dtimeFile,mtimeFile),
                                                              int(prompt(dHours, mHours)),
                                                              prompt(dtimeName, mtimeName),
                                                              prompt(demailTemp, memailTemp))

  # pdb.set_trace()
  # read emails into classes
  REUGroup = loadStudentInfo(mailFile, Group(eFile=timeFile, eSheet=timeSheet))

  # read student times within the written range
  dutySet = REUGroup.upcomingMembersDuties(time=datetime.datetime.now(), 
                                           timeRange=datetime.timedelta(hours=tHours))

  #open connection to mailServer
  try:
    server = MailServer(SMTP_SERVER)

    # create email and send out for each person
    template = open(emailTemp).read()
    template = template.replace("[from]", FROM).replace("[hours]",str(tHours))
    for (name, duties) in dutySet.items():
      # get required information and check for existense of member
      if REUGroup.isMember(name):
        fullName = REUGroup.fullName(name)
        email = REUGroup.members[fullName].email
        msg = server.createMessage(name,duties, template)
  
        # send it out
        server.sendemail(msg,email, From, SUBJECT)

        # tell the user
        print("Sent email to {}({}) for {} upcoming shifts.".format(name, email, len(duties)))
      else:
        print("Error. Attempted to send email to {}({}), but they are not in the group.".format(
              name,fullName))
        print(duties)

  finally:
    # close connection
    server.quit()

if __name__ == '__main__':
  main()
