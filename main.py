from __future__ import print_function

import csv
import pytz
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
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', '1BREK-UFOkW4Fxa0Mb4JIeNFCSdIyXHUSJWHnVOsqDiE')
WORKSHEET_NAME = os.getenv('WORKSHEET_NAME', 'rows')
time_zone = os.getenv('TIME_ZONE', 'Asia/Taipei')
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
        calendars = [
            'lee.shiueh@flh.com.tw'
            # 'gary.li@flh.com.tw',
            # 'benson.lee@flh.com.tw',
            # 'emerson.hsiao@flh.com.tw',
            # 'mark.chen@flh.com.tw'
        ]
        rows = [['status', 'summary', 'tag', 'gw', 'hour', 'year', 'week', 'start', 'end',
                 'calendar', 'organizer', 'id']]
        for calendar_id in calendars:
            print(f'get calendar {calendar_id} event ...')
            events = get_calendar_events(service=service,
                                         calendar_id=calendar_id,
                                         date_since=datetime.date(2021, 1, 1))

            if events:
                rows += events
        with open(f'output.csv', 'w') as fh:
            writer = csv.writer(fh)
            writer.writerows(rows)

        service = build('sheets', 'v4', credentials=creds)
        spreadsheets = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID, fields='sheets.properties'
        ).execute().get('sheets')
        sheet_id = None
        for sheet in spreadsheets:
            if 'title' in sheet['properties'].keys():
                if sheet['properties']['title'] == WORKSHEET_NAME:
                    sheet_id = sheet['properties']['sheetId']
                    break
        if not sheet_id:
            print('Spreadsheets not found')
            return

        with open('output.csv', 'r') as fh:
            content = fh.read()
        body = {
            'requests': [{
                'pasteData': {
                    "coordinate": {
                        "sheetId": sheet_id,
                        "rowIndex": "0",  # adapt this if you need different positioning
                        "columnIndex": "0",  # adapt this if you need different positioning
                    },
                    "data": content,
                    "type": 'PASTE_NORMAL',
                    "delimiter": ',',
                }
            }]
        }
        update_result = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body)

    except HttpError as error:
        print('An error occurred: %s' % error)


def _parsing_events_result(events_result, calendar_id):
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
        return []

    # Prints the start and name of the next 10 events
    rows = []
    for event in events:
        event_id = event['id']
        event_status = event['status']
        summary = event.get('summary', '').replace(' ', '')
        # tag
        if summary.find('場勘') >= 0:
            tag = '場勘'
        elif summary.find('驗收') >= 0:
            tag = '驗收'
        elif summary.find('施工') >= 0:
            tag = '施工'
        elif summary.find('維修') >= 0:
            tag = '維修'
        elif summary.find('測試') >= 0:
            tag = '測試'
        else:
            tag = 'Other'

        # gw
        gw = ''
        if tag == '施工':
            if summary.find('[F') == 0:
                if summary.find('[Ferqo') == 0:
                    gw = 'Ferqo'
                else:
                    gw = 'Fibaro'
            elif summary.find('[X') == 0:
                gw = 'SmartX'
            else:
                gw = ''

        organizer = event.get('organizer', {}).get('email', '')
        _start = event['start'].get('dateTime', event['start'].get('date'))
        start = datetime.datetime.fromisoformat(_start)
        yr = start.year
        wk = start.isocalendar().week
        _end = event['end'].get('dateTime', event['end'].get('date'))
        end = datetime.datetime.fromisoformat(_end)
        date_diff = end - start
        hours = date_diff.seconds / 3600
        # start = start.strftime('%Y-%m-%d %H:%M:%S')
        # end = end.strftime('%Y-%m-%d %H:%M:%S')
        start = start.isoformat()
        end = end.isoformat()
        data = [event_status, summary, tag, gw, hours, yr, wk, start, end, calendar_id, organizer, event_id]
        rows.append(data)
    return rows


def get_calendar_events(service, calendar_id='primary', date_since=None, tz_info=pytz.timezone(time_zone)):
    if not isinstance(date_since, datetime.date):
        date_since = datetime.date.today() + datetime.timedelta(days=-1)

    time_min = datetime.datetime.combine(date_since, datetime.datetime.min.time())
    time_min = time_min.replace(tzinfo=tz_info).isoformat()
    time_max = datetime.datetime.combine(datetime.date.today(), datetime.datetime.max.time())
    time_max = time_max.replace(tzinfo=tz_info).isoformat()
    # print(time_min, time_max)
    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=time_min,
        timeMax=time_max,
        singleEvents=True,
        orderBy='startTime').execute()
    rows = _parsing_events_result(events_result=events_result, calendar_id=calendar_id)
    # print(f'get init rows count {len(rows)} for calendar_id')

    while events_result.get('nextPageToken'):
        next_page_token = events_result.get('nextPageToken')
        events_result = service.events().list(
            calendarId=calendar_id,
            pageToken=next_page_token,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime').execute()
        rows += _parsing_events_result(events_result=events_result, calendar_id=calendar_id)
        # print(f'get page rows count {len(rows)} for calendar_id')

    print(f'get total rows {len(rows)} for {calendar_id}')
    return rows


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
    print(f'now {now}')
    the_date = datetime.date.today() + datetime.timedelta(days=-1)
    time_min = datetime.datetime.combine(the_date, datetime.datetime.min.time())
    time_min = time_min.replace(tzinfo=pytz.timezone('Asia/Taipei')).isoformat()
    print(f'time_min {time_min}')

    print('Getting the upcoming 10 events')
    events_result = service.events().list(calendarId='primary', timeMin=time_min,
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
