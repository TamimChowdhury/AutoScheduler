from __future__ import print_function
import PyPDF2
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import tkinter as tk
from tkinter import filedialog


# Google Calendar API QuickStart

SCOPES = ['https://www.googleapis.com/auth/calendar']

def insertCal(eventName, startTime, endTime):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """

    # Create Event Paramaters

    workEvent = {
    'start': {
    'dateTime': startTime,
    'timeZone': 'America/Toronto',
    },
    'end': {
    'dateTime': endTime,
    'timeZone': 'America/Toronto',
    },
    'summary': eventName
    }

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '/Users/tamimchh/Desktop/Programming-Projects/Python/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # pylint: disable=E1101
    service = build('calendar', 'v3', credentials=creds)
    event = service.events().insert(calendarId='primary', body=workEvent).execute()


# Prompt User to input schedule.pdf via GUI

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()

# Accessing schedule.pdf and assigning all text to variable scheduleText

pdfFileObj = open(file_path, 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
pageObj = pdfReader.getPage(0)
scheduleText = pageObj.extractText().split()

# Idea: Collect shifts that I am working and put them into an array

supervisors = ["KOATHA", "LE-ANN", "MAEGAN", "TAMIM", "PATRICIA", "HANNAH", "ZOE"]
week1schedule = []
week2schedule = []
week3schedule = []
week4schedule = []
scheduler = [week1schedule, week2schedule, week3schedule, week4schedule]
tamimLine = False
daySched = ""
AMPMCounter = 0
weekChoice = 0


for word in scheduleText:
    if (tamimLine == True):
        if word in supervisors or daySched in supervisors:
            tamimLine = False
            daySched = ""
        elif (word == "X" or "VAC" in word or "STAT" in word):
            daySched = word
            scheduler[weekChoice].append(daySched)
            daySched = ""
        elif ("AM" in word or "PM" in word or "CL" in word):
            daySched = daySched + word
            AMPMCounter = AMPMCounter + 1
            if (AMPMCounter == 2):
                scheduler[weekChoice].append(daySched)
                AMPMCounter = 0
                daySched = ""
        else:
            daySched = daySched + word

    
    if (len(scheduler[weekChoice]) == 7):
        weekChoice = weekChoice + 1
    
    if weekChoice > 3:
        break

    else:
        if (word == "TAMIM"):
            tamimLine = True

# Collect all dates and put them into another array:

years = ["2019", "2020"]
months = ["january", "feburary", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
date = ""
week1dates = []
week2dates = []
week3dates = []
week4dates = []
datescheduler = [week1dates, week2dates, week3dates, week4dates]
dateCounter = 0
dateChoice = 0

for word in scheduleText:
    if word.lower() in months:
        date = word
        dateCounter = 3
    elif dateCounter > 1:
        date = date + " " + word
        dateCounter = dateCounter - 1
    
    if dateCounter == 1:
        datescheduler[dateChoice].append(date)
        date = ""
        dateCounter = 0
    
    if (len(datescheduler[dateChoice]) == 7):
        dateChoice = dateChoice + 1
    
    if dateChoice > 3:
        break

# Convert lists to a dictionary

w1dic = dict(zip(week1dates, week1schedule))
w2dic = dict(zip(week2dates, week2schedule))
w3dic = dict(zip(week3dates, week3schedule))
w4dic = dict(zip(week4dates, week4schedule))
shiftlist = []
shiftlist.append(w1dic)
shiftlist.append(w2dic)
shiftlist.append(w3dic)
shiftlist.append(w4dic)

# Fix dates in dictionary

shiftStarts = ""
shiftEnds = ""
startTimes = []
endTimes = []

for dictionary in shiftlist:
    for key, value in dictionary.items():
        shift = key + "!!" + value
        if not "X" in (shift) and not "VAC" in (shift) and not "STAT" in (shift):
        
            shiftDate = shift.split("!!")[0]
            ShiftStartHour = value.split("-")[0]
            ShiftEndHour = value.split("-")[1]
            shiftStartTime = shiftDate + ShiftStartHour
            shiftEndTime = shiftDate + ShiftEndHour
            
            if not ":" in shiftStartTime:
                shiftStarts = datetime.datetime.strptime(shiftStartTime,"%B %d, %Y%I%p")
                startTimes.append(shiftStarts)
            else:
                shiftStarts = datetime.datetime.strptime(shiftStartTime,"%B %d, %Y%I:%M%p")
                startTimes.append(shiftStarts)
            
            if not ":" in shiftEndTime and not "CL" in shiftEndTime:
                shiftEnds = datetime.datetime.strptime(shiftEndTime,"%B %d, %Y%I%p")
                endTimes.append(shiftEnds)
            elif "CL" in shiftEndTime:
                shiftEndTime = shiftEndTime.replace("CL", "11:59PM")
                shiftEnds = datetime.datetime.strptime(shiftEndTime,"%B %d, %Y%I:%M%p")
                endTimes.append(shiftEnds)
            else:
                shiftEnds = datetime.datetime.strptime(shiftEndTime,"%B %d, %Y%I:%M%p")
                endTimes.append(shiftEnds)

# Call Google API to add shifts to the Google Calendar (which is synced w Apple Calendar)

for st, et in zip(startTimes, endTimes):
    insertCal("Working (CN Tower)", st.isoformat('T') + "-05:00", et.isoformat('T') + "-05:00")
    print("Adding Shift on " + st.isoformat() + " to Google Calendar")
    print("Writing Shift to " + st.isoformat() + " to Excel PaySheet")