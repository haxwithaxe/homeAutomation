#!/usr/bin/python
#
########################################################################
# author: chris koepke haxwithaxe@gmail.com
# last update: 21/04/2010
# licence: GPL http://www.gnu.org/licenses/gpl.html
# prerequisites: python-gdata, python-festival, python-feedparser,
#                python-atom, a running festival install,
#                [a google acount with the todo spreadsheet],
#                [the same with a calendar]
#
# description: this script will read the current weather the
#     forcast(US only), weather warnings(US only), any rss feed you want,
#     and also a todo list based on this template
#     https://spreadsheets.google.com/ccc?key=0Ajdl1XiKQer6dFFpWDg0ODMteHFtV3ZycmU2TTh6X1E&hl=en
#     and it will anounce google calendar events for the day.
#
# notes: place your google usename and password in the user and pw
# fields inorder totake advantage of the google docs and google calendar
# functionality
########################################################################

try:
  from xml.etree import ElementTree
except ImportError:
  from elementtree import ElementTree
import xml.dom.minidom
import feedparser
import festival
import datetime
import time
import re
import gdata.spreadsheet.service
import gdata.calendar.service
import gdata.service
import atom.service
import gdata.spreadsheet
import gdata.calendar
import atom
import getopt
import sys
import urllib
import string
from operator import itemgetter

d = datetime.datetime
user = 'googleuser'
pw = 'googlepassword'

# this one is for DC metro uses you can change this to something else if you want
metrourl = 'http://www.wmata.com/rider_tools/metro_service_status/feeds/rail.xml'
# this one is for MARC Penn line users change to another rss if you want
marcurl = 'http://alerts.marylandmail.com/feeds/alerts/2/'
# these two are for the DC Baltimore coridor
warn_md004_url = 'http://weather.noaa.gov/pub/data/watches_warnings/special_weather_stmt/md/mdz004.txt'
warn_md003_url = 'http://weather.noaa.gov/pub/data/watches_warnings/special_weather_stmt/md/mdz003.txt'

# DEFAULT SPREADSHEET TITLE!!!
HOMETODO = 'Home Todo'
# alternate spreadsheet title
WORKTODO = 'Work Todo'

# SCHEDULING ###########################################################
# THESE MUST BE IN THE TREE LETTER DAY OF THE WEEK FORMAT
workdays = ['Tue','Thu'] # work days
offdays = ['Mon','Wed','Fri','Sat','Sun'] # days off

# exception types: single day, recuring, and messages (the computer will read the message on the date of the exception)
exceptions = ['16/02/2010'] # this is a list of dates that will either be silent or will have a special message set below
recurring = ['25/12'] # these are recurring dates that have exceptions associated with them year to year
# this is a dictionary of dates that have messages associated with them
messages = {'12/12/2012': 'Welcome to the end of the world', '16/02/2010': 'This is a test of the exceptions message functionality'}
########################################################################

## regex stuff do not touch!!
dateformat = '%d/%m/%Y' # day/month/year -- 01/01/2070
recurringdateformat = '%d/%m' # day/month
weekdayformat = '%A' # you probly don't want to change this it will change what the voice says
NWSDateFormat = 'Expires:%Y%m%d%H%M'
brre = re.compile('<br />')
slashre = re.compile('/')
warnDatere = re.compile('Expires:[0-9]{12}')
warnMsgre = re.compile('[0-9]{3,4} [AP]M [A-Z]{3} [A-Z]{3} [A-Z]{3} [0-9]{2} [0-9]{4}[\n]\.\.\.[A-Z]*[\S\s\n\r]*\$\$')

GOOGLETSFORMAT = '%Y-%m-%dT%H:%M:%S'
CALSAYDATE = '%H:%M%p'
TODAY = d.today()
TOMORROW = TODAY + datetime.timedelta(days=1)

class getGcalItems:

   def __init__(self, email, password):
      self.cal_client = gdata.calendar.service.CalendarService()
      self.cal_client.email = email
      self.cal_client.password = password
      self.cal_client.source = 'Google-Calendar_Python_Sample-1.0'
      self.cal_client.ProgrammaticLogin()
      self.author = 'api.rboyd@gmail.com (Ryan Boyd)'

   def _DateRangeQuery(self, start_date='1970-01-01', end_date='2069-12-31'):
      cal_list = self.cal_client.GetAllCalendarsFeed()
      for c in cal_list.entry:
         uri = urllib.unquote(c.id.text.split('/')[8])
         #print(uri)
         query = gdata.calendar.service.CalendarEventQuery(uri,'private','full')
         query.start_min = start_date
         query.start_max = end_date 
         feed = self.cal_client.CalendarQuery(query)
         for i in feed.entry:
            who = []
            if i.title.text:
               string = i.title.text
            else:
               string = 'Appointment'
            for w in i.who:
               if w.name != feed.title.text:
                 who.append(w.name)
            try:
               if len(who) > 0:
                  string += ' with '
                  if len(who) > 1:
                     for l,p in enumerate(who):
                        if l == 0:
                           string += p
                        else:
                           string += ' and '+p
                  else:
                     string += 'with '+who[0]
            except:
               if len(i.content.text) > 0:
                  string += 'with '+i.content.text
            if i.where[0].value_string:
               string += ' at '+i.where[0].value_string
            if i.when[0].start_time:
               try:
                  string += ' at '+d.strptime(i.when[0].start_time.split('.')[0],GOOGLETSFORMAT).strftime(CALSAYDATE)
               except:
                  continue
         festival.say(string+' ...')

   def Run(self,start_date,end_date):
      return self._DateRangeQuery(start_date,end_date)

class getTodoItems:
  def __init__(self, email, password):
    self.gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    self.gd_client.email = email
    self.gd_client.password = password
    self.gd_client.source = 'Spreadsheets GData Sample'
    self.gd_client.ProgrammaticLogin()
    self.curr_key = ''
    self.curr_wksht_id = ''
    self.list_feed = None
    self.author = 'api.laurabeth@gmail.com (Laura Beth Lincoln)'

  def _ListGetAction(self):
    # Get the list feed
    list_feed = self.gd_client.GetListFeed(self.curr_key, self.curr_wksht_id)
    tlist = []
    for i in list_feed.entry:
      item = (i.title.text,i.custom['due'].text,i.custom['pri'].text,i.custom['status'].text,i.custom['comments'].text)
      tlist.append(item)
    return tlist

  def _StringToDictionary(self, row_data):
    dict = {}
    for param in row_data.split():
      temp = param.split('=')
      dict[temp[0]] = temp[1]
    return dict

  def Run(self, SpreadSheetID, WorkSheetID = '0'):
    # get real spreadsheet ID
    feed = self.gd_client.GetSpreadsheetsFeed()
    for i in feed.entry:
      if i.title.text == SpreadSheetID:
         id_parts = i.id.text.split('/')
         break

    self.curr_key = id_parts[len(id_parts) - 1]
    # get real worksheet ID
    feed = self.gd_client.GetWorksheetsFeed(self.curr_key)
    self.sheet_title = feed.title.text
    id_parts = feed.entry[string.atoi(WorkSheetID)].id.text.split('/')
    self.curr_wksht_id = id_parts[len(id_parts) - 1]
    return self._ListGetAction()

def unabriv(strin):
   strstuff = strin.replace('MPH','miles per hour').replace(' N ','north').replace('NNE','north north east').replace('NE','north east').replace('ENE','east north east').replace(' E ',' east ').replace('ESE','east south east').replace('SE','south east').replace('SSE','south south east').replace(' S ',' south ').replace('SSW','south south west').replace('SW','south west').replace('WSW','west south west').replace(' W ',' west ').replace('WNW','west north west').replace('NW','north west').replace('NNW','north north west').replace('MD','maryland').replace('VA','virginia').replace('INTL','international').replace('WMATA','W M A T A').replace(' F ','degrees fahrenheit').replace(' C)',' degrees celsius').replace('&amp;','and')
   strout = strstuff #parenre.sub('',strstuff)
   return strout

def get_url(url):
   """Return a string containing the results of a URL GET."""
   import urllib2
   try: return urllib2.urlopen(url).read()
   except urllib2.URLError:
      festival.say('I failed to retrieve the required data ...')
      return False

def get_metar(Id, headers=None, murl=None):
   """Return a summarized METAR for the specified station."""
   if not murl:
      murl = "http://weather.noaa.gov/pub/data/observations/metar/decoded/%ID%.TXT"
   murl = murl.replace("%ID%", Id.upper())
   murl = murl.replace("%Id%", Id.capitalize())
   murl = murl.replace("%iD%", Id)
   murl = murl.replace("%id%", Id.lower())
   murl = murl.replace(" ", "_")
   metar = get_url(murl)
   if not metar:
      print('no metar data was returned from the url provided\n')
      return False
   lines = metar.split("\n")
   if not headers:
      headers = \
         "relative_humidity," \
         + "precipitation_last_hour," \
         + "sky conditions," \
         + "temperature," \
         + "weather," \
         + "wind"
   headerlist = headers.lower().replace("_"," ").split(",")
   for header in headerlist:
      for line in lines:
         if line.lower().startswith(header + ":"):
            if line.endswith(":0"):
               line = line[:-2]
            msg = unabriv(line)+' ...'
            festival.say(msg)


def get_forecast(city, st, flines="0", furl=None):
   """Return the forecast for a specified city/st combination."""
   if not furl:
      furl = "http://weather.noaa.gov/pub/data/forecasts/city/%st%/%city%.txt"
   furl = furl.replace("%CITY%", city.upper())
   furl = furl.replace("%City%", city.capitalize())
   furl = furl.replace("%citY%", city)
   furl = furl.replace("%city%", city.lower())
   furl = furl.replace("%ST%", st.upper())
   furl = furl.replace("%St%", st.capitalize())
   furl = furl.replace("%sT%", st)
   furl = furl.replace("%st%", st.lower())
   furl = furl.replace(" ", "_")
   forecast = get_url(furl)
   if not forecast:
      print('no forcast returned from the url provided')
      return False
   lines = forecast.split("\n")
   if not flines: flines = len(lines) - 5
   for line in lines[5:10]:
      if line.startswith("."):
         msg = unabriv(line).replace(".", "", 1)
         festival.say(msg)


def get_weather(Id, city, st):
   get_metar(Id)
   get_forecast(city,st)

def get_warning(url):
   warnings = get_url(url)
   if not warnings:
      print('no data returned from url provided\n')
      return False
   try:
      date = warnDatere.findall(warnings)
   except:
      print('no date in page from url\n')
      return False
   if d.strptime(date[0],NWSDateFormat) > d.today():
      msgs = warnMsgre.findall(warnings)
      for m in msgs:
         msg = unabriv(m)+' ... ...'
         festival.say('Weather Advisory ...')
         festival.say(msg)
   else:
      print('expired message\n')
   return True

def sayRSS(url):
   try:
      f = feedparser.parse(url)
   except:
      festival.say('I failed to retrieve the required data ...')
      return 1
   festival.say(f.feed.title)
   for i in f.entries:
      msg = unabriv(re.sub(brre,'\n',i.description).replace('/',' or '))+' ...'
      festival.say(msg)
      print(msg)

def say_todos(ToDos):
   sortfirst = 1
   sortsecond = 2
   todosheet = getTodoItems(user, pw)
   s = sorted(todosheet.Run(ToDos),key=itemgetter(sortfirst))
   todos = sorted(s,key=itemgetter(sortsecond))
   todolist = []
   lowprilist = []
   for i in todos:
      if i[2] != '0' and i[2] != None:
         output = i[0]
         if i[1] != None:
            output += ' by '+d.strptime(i[1],'%m/%d/%Y').strftime('%A %d %B')
         output += ' ...'
         todolist.append(output)
      else:
         lowpri = i[0]
         if i[1] != None:
            lowpri += ' by '+d.strptime(i[1],'%m/%d/%Y').strftime('%A %d %B')
         lowpri += ' ...'
         lowprilist.append(lowpri)
   for i in todolist:
      festival.say(i)
   for i in lowprilist:
      festival.say(i)

def say_gcal(when):
   gcalsession = getGcalItems(user, pw)
   if when == 'today':
      gcalsession.Run(TODAY.strftime('%Y-%m-%d'),TOMORROW.strftime('%Y-%m-%d'))

def main():
   saytodo = True
   todolist = HOMETODO
   now = datetime.date.today()
   today = now.strftime('%a')
   thisday = 'anyday'

   # the fist thing the computer says
   festival.say('good morning ... it\'s time to get up ...')


   '''for i in workdays:
      if today == i:
         thisday = 'wokday'

   for i in offdays:
      if today == i:
         thisday = 'offday'

   for i in exceptions:
      if now.strftime(dateformat) == i:
         thisday == 'exception'

   if thisday == 'workday':
     # make the computer say the weather and the date
      festival.say('Transit and Weather Information for '+d.now().strftime('%A %d %B %Y')+' ...')
      #print("It's "+today)
      festival.say('Weather for Kensington ...')
      get_weather('KIAD','washington_dulles_intl_airport','va')
      #print('got the weather for kensington')
      festival.say('Weather for Baltimore ...')
      get_weather('KBWI','baltimore-washington_intl_airport','md')
      #print('got the weather for baltimore')
     # make the computer read the Metro and MARC RSS data
      sayRSS(metrourl)
      sayRSS(marcurl)
      saytodo = True
      todolist = WORKTODO
   if thisday == 'offday':
     # same as above just for my home town instead of both home and school
      festival.say('Transit and Weather Information for '+d.now().strftime('%A %d %B %Y')+' ...')
      festival.say('Weather for Kensington ...')
      get_weather('KIAD','washington_dulles_intl_airport','va')
      saytodo = True
   if thisday in ('workday','offday'):
     #regardless fo the type of day say the weather warnings that apply
      get_warning(warn_md004_url)
      get_warning(warn_md003_url)
   if thisday == 'exception':
      festival.say(messages[now.strftime(dateformat)])'''

   if saytodo == True:
     # say the todo list from google docs spreadsheet
      #festival.say('Your To Do List And Appointments For Today ...')
      #say_todos(todolist)
      say_gcal('today')


if __name__ == '__main__':
  main()
