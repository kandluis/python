#/usr/bin/env python

import smtplib
import re
import csv
import pdb

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# regex for email checks
EMAIL_REGEX = re.compile(r"[^@]+@[^@]+\.[^@]+")

# specifies tolerance for error when matching names
ERR_TOL = 0.10 # value from 1 to 2, with 1 requiring perfect match, and 2 not comparing


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Helper Functions Used in Script - Should be Factored out if using in a larger
project.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
def prompt(default,message):
  '''
  Promts the users with message (no additional characters). If user enters nothing,
  default is returned, otherwise the user input is returns.
  '''
  inp = input(message)
  return (inp if inp != '' else default)

def sendemail(message,To,From,Subject):
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
  server.send_message(msg)

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
      match = 1 if s[i-1] == t[j-1] else 0
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
  correctly.
  '''
  try:
    return s.split(' ')[i-1]
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
  return(re.sub(r'\W+( \W+)*$', '', name))

class Student(object):
  '''
  A student class to keep track of first name, last name, and email for a student
  '''
  def __init__(self,name,email):
    '''
    Student name and email in default input formats. 
    '''
    self.name = name
    self.email = email if EMAIL_REGEX.match(email) else None

class Group(object): 
  ''' 
  A group class keeps track of a group of students, and provides membership and 
  lookup functions for students 
  '''
  def __init__(self,name = "REU"): 
    self.name = name 
    self.size = 0
  
    # keep track of student objects by full names 
    self.members = {} 
  
    # keep track of names of student objects 
    self.memberNames = []  
  
  def addMember(self,member): 
    ''' 
    Creates a link between fullname and the student object. Returns True if  
    succesfull, False if student already exists. Cleans the name.  
    '''
    if member.name not in self.members and member.name not in self.memberNames: 
      self.members[cleanName(member.name)] = member 
      self.memberNames.append(member.name)
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
    # clean name 
    cName = cleanName(name) 
      
    # returns list of matched items based on index 
    def matchList(i): 
      if i == 0: 
        return list(map(lambda t : stringMatchDec(cName,t), self.memberNames)) 
      else: 
        return list(map(lambda t : stringMatchDec(ithWord(cName,i), t),  
                        map(lambda s: ithWord(s,i), self.memberNames))) 
  
    # match fullnames 
    flNameMatchDec = matchList(0) 
    flNameInd = minIndex(flNameMatchDec) 
    if flNameMatchDec[flNameInd] <= ERR_TOL: 
      return self.memberNames[flNameInd] 
  
    # match firstnames 
    fNameMatchDec = [full + first for (full, first) in zip(flNameMatchDec,  
                                                           matchList(1))]  
    fNameInd = minIndex(fNameMatchDec) 
    if fNameMatchDec[fNameInd] <= 2*ERR_TOL: 
      return self.memberNames[fNameInd] 
  
    # match last names 
    lNameMatchDec = [prev + last for (prev,last) in zip(fNameMatchDec,
                                                        matchList(-1))] 
    lNameInd = minIndex(lNameMatchDec) 
    if lNameMatchDec[lNameInd] <= 3*ERR_TOL:
      return self.memberNames[lNameInd]
 
    # nothing matched
    return None

  def findMemberr(self,name): 
    ''' 
    Finds a member and returns the member object if found. Otherwise, returns None 
    '''
    if self.isMember(name): 
      return self.members[self.fullName(name)] 
    else: 
      return None
      
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
    
if __name__ == '__main__':
  # defaults
  dFrom, dmailFile, dtimeFile = "sknapp@fas.harvard.edu", "REU Student Emails.csv","Meal Shifts 2014.xlsx"

  # messages
  mFrom = "From email? ({}): ".format(dFrom)
  mmailFile = "Email Address File ({}): ".format(dmailFile)
  mtimeFile =  "Time Sheet File ({}): ".format(dtimeFile)

  # prompt for modifications
  From, mailFile, timeFile = prompt(dFrom,mFrom), prompt(dmailFile, mmailFile), prompt(dtimeFile,mtimeFile)

  # read emails into classes
  REUGroup = loadStudentInfo(mailFile, Group())

  # read student times


  #open connection to mailServer
  #server = smtplib.SMTP('fas.harvard.edu')




  # close connection
  #server.quit()
