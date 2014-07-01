import re
import datetime

class Duty(object):
  '''
  Contains information on a single type of Meal Duty
  '''
  def __init__(self,rows,meal,mtype, timeString):
    self.rows = rows
    self.meal = meal
    self.type = mtype
    self.time = timeString

  def name(self):
    '''
    Returns a human readable description of the Duty
    '''
    return(self.meal + " " + self.type)

# maps times to specific columns in the excel timesheet data as well as descriptors
TIMECOLS = {  "7:00AM"  : Duty([1], "Breakfast", "Prep.", "7:00 AM"),
              "8:00AM"  : Duty([2], "Breakfast", "Clean-Up", "8:00 AM"),
              "11:30AM" : Duty([3], "Lunch", "Prep.", "11:30 AM"),
              "12:30PM" : Duty([4,5], "Lunch", "Clean-Up", "12:30 PM"),
              "5:00PM"  : Duty([6], "Dinner", "Prep.", "5:00 PM"),
              "6:00PM"  : Duty([7,8], "Dinner", "Clean-Up", "6:00 PM")
}

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Helper Functions Used in MealReminder Program
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
def checkTime(limBelow, limAbove):
  '''
  Returns the list of (index, duty) of the columns whose time is limBelow <= time <= limAbove
  @limAbove and limBelow are both (hh, mm) tuple.
  '''
  def underLim(t):
    '''
    etunrs true if the time t (which is a string in format HH:MM{PM/AM}) occurs
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

def numWords(s):
  return(len(s.split(" ")))

def cleanName(name, upper=True): 
  '''
  Removes all non-characters from name
  '''
  # remove extraneous symbols and upper
  cName = re.sub(r'\W+( \W+)*$', '', name).upper() if upper else re.sub(r'\W+( \W+)*$', '', name)

  # some names have an extra space at the end
  try:
    if cName[-1] == " ":
      return(cName[0:len(cName)-1])
    return(cName)
  except IndexError:
    return(cName)
