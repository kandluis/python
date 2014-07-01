import smtplib
import re
import xlrd
import datetime
from email.mime.text import MIMEText
from includes.functions import *

# regex for email checks
EMAIL_REGEX = re.compile(r"[^@]+@[^@]+\.[^@]+")
# regex for dates of the form mm/dd
DATE_REGEX = re.compile(r"^(0?[1-9]|1[0-2])/(0?[1-9]|[12][0-9]|3[01])$")


# specifies tolerance for error when matching names
ERR_TOL = 0.50 # value from 1 to 2, with 1 requiring perfect match, and 2 not comparing
# contants for the date,row tuple we use
DATE = 0
ROW = 1

EARLIEST_TIME = (7,0)
LATEST_TIME = (18,0)

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
      dStr += "On " + date + " at " + duty.time + ": " + duty.meal + " " + duty.type + "\n\n"

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
        if DATE_REGEX.match(str(dateStr)):
          month, day = dateStr.split("/")
          E_hr, E_mm = EARLIEST_TIME
          date = datetime.datetime(year=time.year, month=int(month), day=int(day),
                                   hour=E_hr, minute=E_mm)
          initDate = datetime.datetime(year=time.year, month=time.month, day=time.day,
                                       hour=E_hr, minute=E_mm) - datetime.timedelta(days=1)
          return(initDate < date < (time + timeRange))
        else:
          return False
      
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
      
      def add(d,cell,v):
        # try to split the cell input so we can handle multiple names
        nameList = [cleanName(name, upper=False) for name in cell.split("  ") if accept(x)]
        for name in nameList:
          if name in d:
            d[name].append(v)
          else:
            d[name] = [v]

      def isWorkDay(dayStr):
        '''
        Input is of format mm/dd where mm/dd is a workday. Compare current day/month
        to it and return true if the two match.
        '''
        try:
          if DATE_REGEX.match(str(dayStr)):
            month, day = dayStr.split("/")
            nmonth, nday = time.month, time.day
            return(str(nmonth) == month and str(nday) == day)
          else:
            return False
        except (ValueError, AttributeError) as e:
          return False


      # get limits
      lim = time + timeRange
      limBelow = (time.hour, time.minute)
      limAbove = (lim.hour, lim.minute) if lim.hour > time.hour else LATEST_TIME 

      # create the dictionary to hold 
      matchings = {}

      if len(rows) == 1:
        for (i,duty) in checkTime(limBelow, limAbove):
          add(matchings,rows[0][ROW][i], (rows[0][DATE],duty))
      # multiple
      elif len(rows) > 1:
        # first row (limit further if today is a work day)
        firstDay = limBelow if isWorkDay(rows[0][DATE]) else EARLIEST_TIME
        for (i,duty) in checkTime(firstDay, LATEST_TIME):
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
