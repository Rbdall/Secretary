'''
Created on Sep 11, 2013

@author: Ryan Dall
'''

try:
    from xml.etree import ElementTree
except ImportError:
    from elementtree import ElementTree
    
import gdata.calendar.data
import gdata.calendar.client
import atom.data
import gspread
import time

#Credentials of Psi/EK calendar account
user = "###"
password = "###"

#URL of google form response spreadsheet
eventURL = "###"

#Attempts to connect to google with given credentials
def attemptLogin():
    client = gdata.calendar.client.CalendarClient(source='Psi/EK-Calendar-v1')
    client.ClientLogin(user, password, client.source)
    return(client)
    

#Returns an array of titles of all events on the calendar which have yet to pass       
def getAllCurrentEvents(calendar_client):
    currentEvents = []
    query = gdata.calendar.client.CalendarEventQuery()
    query.start_min = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())
    feed = calendar_client.GetCalendarEventFeed(q=query)
    for i, an_event in enumerate(feed.entry):
        currentEvents.append(an_event.title.text)
    return currentEvents

#Returns a list of tuples. Each tuple consists of an event and a boolean marking new events vs edits       
def getAllSpreadsheetEvents():
    events = []
    login = gspread.login(user, password)
    sheet = login.open_by_url(eventURL)
    worksheet = sheet.get_worksheet(0)
    for x in range(2, worksheet.row_count):
        info = worksheet.row_values(x)
        if (info == []):
            return events
        events.append(scrapeEvent(info))

        
    return events

#Creates a single event from a row 
def scrapeEvent(row):
    eventEntry = []
    
    #Pulling pertinent data from the spread sheet. Optional fields are protected by try/catch blocks
    submitter = row[1] + " " + row[2]
    
    editing = False
    if row[3] is not None:
        editing = True 
    title = row[4]
    
    try:
        where = row[5]
    except:
        where = ""
        
    date = row[6].split('/')
    start = row[7].split(':')
    end = row[8].split(':')
    
    try:
        organizer = row[9]
    except:
        organizer = ""
    
    try:
        desc = row[10]
    except:
        desc = ""
        
    
    #Create event and append text fields
    event = gdata.calendar.data.CalendarEventEntry()
    event.title = atom.data.Title(text=title)
    event.content = atom.data.Content(text='Submitted by: ' + submitter + '\n' 
                                      + "Please direct questions to: " + organizer + '\n' + '\n'
                                      + "Additional information: " + desc)
    event.where.append(gdata.calendar.data.CalendarWhere(value=where))
    
    #Process dates into rfc3339 format and append
    if int(date[1]) < 10:
        date[1] = "0" + str(date[1])
    if int(date[0]) < 10:
        date[0] = "0" +str(date[0])
    if int(start[0]) < 10:
        start[0] = "0" + str(start[0])
    if int(end[0]) < 10:
        end[0] = "0" + str(end[0])
    start_time = date[2]+"-"+date[0]+"-"+date[1]+"T"+start[0]+":"+start[1]+":"+start[2]+"-07:00"
    end_time = date[2]+"-"+date[0]+"-"+date[1]+"T"+end[0]+":"+end[1]+":"+end[2]+"-07:00"
    event.when.append(gdata.calendar.data.When(start=start_time, end=end_time))
    
    
    
    if(time.strptime(date[2] + " " + date [0] + " " + date[1] + " " + 
                     start[0] + " " + start[1] + " " + start[2], '%Y %m %d %H %M %S') < time.localtime()):
        eventEntry.append(None)
    else:
        eventEntry.append(event)
    eventEntry.append(editing)
    return eventEntry
    
    
#Adds event to the calendar        
def InsertSingleEvent(calendar_client, event):
    new_event = calendar_client.InsertEvent(event)

    print 'New single event inserted: %s' % (new_event.title.text,)


    return new_event

#Edits an existing event on the calendar
def EditSingleEvent(calendar_client, eventToChange, newEvent):
    eventToChange.title = newEvent.title
    eventToChange.content = newEvent.content
    eventToChange.where = newEvent.where
    for a_when in newEvent.when:
        startTime = a_when.start
        endTime = a_when.end
    eventToChange.when[0].start = startTime
    eventToChange.when[0].end = endTime

    
    calendar_client.Update(eventToChange)
    print 'Event edited: ' + eventToChange.title.text
    
    
if __name__ == '__main__':
    #Attempts to log into google
    client = attemptLogin()
    
    #Pulls list of events to add and events currently on calendar
    eventList = getAllSpreadsheetEvents()
    currentList = getAllCurrentEvents(client)


    #For every event to add
    for events in eventList:
        #If event happened in the past, it will be NoneType
        if(events[0] == None):
            continue
        #If the event already exists and doesn't need to be edited, continue
        if(events[0].title.text in currentList and events[1] == False):
            continue
        #If the event exists and needs to be edited, query for the existing event and update it
        elif(events[0].title.text in currentList and events[1] == True):
            query = gdata.calendar.client.CalendarEventQuery(text_query=events[0].title.text)
            feed = client.GetCalendarEventFeed(q=query)
            for i, change in enumerate(feed.entry):
                EditSingleEvent(client, change, events[0])
                break
        #Else, the event is added to the calendar and the current list of events
        else:
            InsertSingleEvent(client, events[0])
            currentList.append(events[0].title.text)
    

    
    