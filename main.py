from __future__ import print_function

import dotenv
import datetime
import os.path
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
dotenv.load_dotenv()
credential = os.getenv('GOOGLE_API_CREDENTIAL_FILE')
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credential, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)

        # get_upcoming_10_events(service)
        # get_calendar_list(service)
        get_daily_events(service)

    except HttpError as error:
        print('An error occurred: %s' % error)


def get_daily_events(service, calendar_id='primary', the_date=None):
    if not isinstance(the_date, datetime.datetime):
        the_date = datetime.datetime.now() - datetime.timedelta(days=1)
    timeMin = the_date.date().utcnow().isoformat() + 'Z'
    timeMax = (the_date + datetime.timedelta(days=1)).date().utcnow().isoformat() + 'Z'
    print(timeMin, timeMax)


def get_calendar_list(service):
    calendar_list_result = service.calendarList().list().execute()
    if not calendar_list_result:
        print('No calendars list found.')
        return
    calendars = calendar_list_result.get('items', [])
    for cal in calendars:
        print(f'id {cal["id"]}, primary {cal.get("primary")}, '
              f'description {cal.get("description")}, summary {cal.get("summary")}')


def get_upcoming_10_events(service):
    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    print('Getting the upcoming 10 events')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          maxResults=10, singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
        return

    # Prints the start and name of the next 10 events
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])


if __name__ == '__main__':
    main()
