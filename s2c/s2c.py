#!/usr/bin/env python
"""This script takes in a formatted Google Sheet with Beaumont ER Residency
monthly schedule and creates Google Calendar events for a specified
resident."""

from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/calendar']

# Set the IDs of the Calendar and Sheet, the range of the Sheet,
# and the person to create shifts for.
CALENDAR_ID = 'il5tc4f32rv6tt6v5fh5hsbc14@group.calendar.google.com'
SPREADSHEET_ID = '16LG3K0OLyI-ub90V8t6Ept_FEq18B5LANofx9Or308Q'
RANGE_NAME = 'Sheet1!A1:I48'
PERSON = 'Perkis'

class Shift:
    """Shift class represents a working shift."""

    def __init__(self):
        self.year = '2020'
        self.title = ''
        self.time = ''
        self.datenum = ''
        self.month = ''

# This helper function allows to handle weeks where the person
# has more than one shift that week
def get_index_positions(list_of_elements, element):
    """Returns the indexes of all occurrences of given element
    in list_of_elements."""
    index_pos_list = []
    index_pos = 0
    while True:
        try:
            # Search for item in list from indexPos to the end of list
            index_pos = list_of_elements.index(element, index_pos)
            # Add the index position in list
            index_pos_list.append(index_pos)
            index_pos += 1
        except ValueError:
            break
    return index_pos_list

def make_shifts(rawdata):
    """Constructs shift objects."""
    current_month = rawdata[0][0]
    current_year = rawdata[0][1]
    person = PERSON
    shift_list = []
    for rowlist in rawdata:
        if rowlist[0] == current_month:
            current_date_index = 0
        if rowlist[0] == '.':
            current_date_index = rawdata.index(rowlist)
        if person in rowlist:
            person_indexes = get_index_positions(rowlist, person)
            for i in person_indexes:
                shift = Shift()
                shift.title = rowlist[0]
                shift.time = rowlist[1]
                shift.datenum = rawdata[current_date_index][i]
                shift.month = current_month
                shift.year = current_year
                shift_list.append(shift)
    return shift_list

def make_times(timestring):
    """Returns the start and end time for a calendar event."""
    if timestring.lower() == '7a-7p':
        start_time = 7
        length_hours = 12
    elif timestring.lower() == '7p-7a':
        start_time = 19
        length_hours = 12
    elif timestring.lower() == '8a-6p':
        start_time = 8
        length_hours = 10
    elif timestring.lower() == '6p-4a':
        start_time = 18
        length_hours = 10
    elif timestring.lower() == '9a-7p':
        start_time = 9
        length_hours = 10
    elif timestring.lower() == '10a-8p':
        start_time = 10
        length_hours = 10
    return start_time, length_hours

def make_events(shift_list, calendar_service):
    """Create calendar events from list of shift objects."""
    for shift in shift_list:
        start_time, length_hours = make_times(shift.time)
        start_date = datetime.datetime(int(shift.year), int(shift.month),
                                       int(shift.datenum), start_time, 0, 0)
        end_date = (start_date + datetime.timedelta(hours=length_hours))
        event = {
            'summary': shift.title,
            'start': {
                'dateTime': start_date.isoformat(),
                'timeZone': 'America/Detroit',
            },
            'end': {
                'dateTime': end_date.isoformat(),
                'timeZone': 'America/Detroit',
            },
        }
        event = calendar_service.events().insert(calendarId=CALENDAR_ID,
                                                 body=event).execute()
        print('Event created: {}'.format(event.get('htmlLink')))

def main():
    """Creates shift Calendar events given a spreadsheet schedule."""
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Build services to Sheets and Calendar
    sheetservice = build('sheets', 'v4', credentials=creds)
    calservice = build('calendar', 'v3', credentials=creds)

    # Call the Sheets API
    sheet = sheetservice.spreadsheets() # pylint: disable=no-member
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    # Get the raw data from the sheet
    values = result.get('values', [])
    # Process the raw data into shift objects
    all_shifts = make_shifts(values)
    # Create the shift Calendar events
    make_events(all_shifts, calservice)

if __name__ == '__main__':
    main()
