import csv
import datetime
from dateutil import tz
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']

# Time zones to convert actual data to ISO time
CENTRAL_TIME = tz.gettz('America/Chicago')

# Time tracking calendar ID stored in txt file
TIME_TRACKING_CAL_ID = 'data/calendar_id.txt'

# CSV file for time tracking
ATIMELOGGER_REPORT = 'data/report-16-10-2021.csv'


def load_report(report_csv):
    '''
    Load ATimeLogger report and parse 'From' and 'To' fields into datetime objects.
    '''
    entries = []
    with open(report_csv) as report_file:
        reader = csv.DictReader(report_file)
        for row in reader:
            # Discard the "summary" rows at the end of the file
            if row['To'] is None:
                continue

            activity = {}
            activity['Activity type'] = row['Activity type']
            for ft in ['To', 'From']:
                time_str = row[ft]
                date_time = datetime.datetime.fromisoformat(time_str)
                dt_central = date_time.astimezone(CENTRAL_TIME)
                activity[ft] = dt_central
            entries.append(activity)
    return entries

def add_calendar_evt(service, calendar_id, name: str, from_dt: datetime.datetime, to_dt: datetime.datetime):
    '''
    Create an event on the Time Tracking calendar with a specified start and end time
    '''
    body = {
        'summary': name,
        'start': {
            'dateTime': from_dt.isoformat()
        },
        'end': {
            'dateTime': to_dt.isoformat()
        }
    }
    return service.events().insert(calendarId=calendar_id, body=body).execute()

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('data/token.json'):
        creds = Credentials.from_authorized_user_file('data/token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'data/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('data/token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)

    # Load name of calendar
    with open(TIME_TRACKING_CAL_ID) as fin:
        calendar_id = fin.read()

    # Load report csv
    activities = load_report(ATIMELOGGER_REPORT)

    # Fetch min/max date of activities (as datetime objects)
    beginning = min(activities, key=lambda a: a['From'])['From']
    end = max(activities, key=lambda a: a['To'])['To']

    # Get all existing calendar events between beginning and end to ensure no
    # duplicate events are created on gcal
    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=beginning.isoformat(),
        timeMax=end.isoformat(),
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    events = events_result.get('items', [])

    # Create an easy lookup to determine if an event is a duplicate:
    # {"summary1": [{'start': <datetime>, 'end': <datetime>}], "summary2":  ...}
    existing = {}
    for evt in events:
        summary = evt['summary']
        new_rec = {}
        for field in ['start', 'end']:
            new_rec[field] = datetime.datetime.fromisoformat(evt[field]['dateTime'])

        evt_list = existing.get(summary, [])
        evt_list.append(new_rec)
        existing[summary] = evt_list

    for activity in activities:
        act_type = activity['Activity type']
        start = activity['From']
        end = activity['To']

        # First, check if already in calendar
        if act_type in existing:
            if start in map(lambda a: a['start'], existing[act_type]) and end in map(lambda a: a['end'], existing[act_type]):
                print('Activity `{}` starting {} already exists, skipping'.format(act_type, start.isoformat()))
                continue

        print('Adding calendar event for {}, starting {}'.format(act_type, start.isoformat()))
        result = add_calendar_evt(service, calendar_id, act_type, start, end)
        print('  -> Added event `{}`, starting {}'.format(result['summary'], result['start']['dateTime']))

if __name__ == '__main__':
    main()