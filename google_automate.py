import os
import csv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
from google.auth.transport.requests import Request


credentials_file = 'C:/Shashwat/Work/Coding/Google Calendar/credentials.json'
token_file = 'C:/Shashwat/Work/Coding/Google Calendar/token.json'
calendar_id = 'singhshashwat08@gmail.com'
csv_file = 'C:/Shashwat/Work/Coding/Google Calendar/Exported Data/events.csv'

# Set up the Google Calendar API
def create_calendar_service():
    credentials = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token_file):
        credentials = Credentials.from_authorized_user_file(token_file)

    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_file, ['https://www.googleapis.com/auth/calendar.readonly']
            )
            credentials = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open(token_file, 'w') as token:
            token.write(credentials.to_json())

    # Build the service
    service = build('calendar', 'v3', credentials=credentials)
    return service

# Export Google Calendar events to a CSV file
def export_calendar_events(service, calendar_id, csv_file):
    present_time = datetime.utcnow().isoformat() + 'Z'
    start_of_year = datetime(datetime.now().year, 1, 1).isoformat() + 'Z'
    max_time = datetime(datetime.now().year + 1, 1, 1).isoformat() + 'Z'

    try:
        events_result = (
            service.events()
            .list(calendarId=calendar_id, timeMin=start_of_year, timeMax=max_time, maxResults=1000, singleEvents=True, orderBy='startTime')
            .execute()
        )
        events = events_result.get('items', [])

        if not events:
            print(f'No events found.')
        else:
            with open(csv_file, 'w', newline='') as csvfile:
                csv_writer = csv.writer(csvfile)
                header = [
                    'Event ID', 'Event Summary', 'Start Time', 'End Time', 'Date', 'Color', 'Description', 'Is Recurring'
                ]
                csv_writer.writerow(header)

                for event in events:
                    if 'recurrence' in event:
                        recurring_events = expand_recurring_event(service, calendar_id, event)
                        for recurring_event in recurring_events:
                            write_event_to_csv(csv_writer, recurring_event, True)
                    else:
                        write_event_to_csv(csv_writer, event, False)

                print(f'Events exported to {csv_file}.')

    except HttpError as e:
        print(f'Error: {e.content}')

# Expand recurring events into individual instances
def expand_recurring_event(service, calendar_id, event):
    recurring_events = []
    try:
        instances = (
            service.events()
            .instances(calendarId=calendar_id, eventId=event['id'])
            .execute()
            .get('items', [])
        )
        recurring_events.extend(instances)
    except HttpError as e:
        print(f'Error expanding recurring event: {e.content}')
    return recurring_events

# Write an event to CSV
# Write an event to CSV
def write_event_to_csv(csv_writer, event, is_recurring):
    event_id = event['id']
    summary = event.get('summary', 'N/A')
    start_time = event['start'].get('dateTime', event['start'].get('date'))
    end_time = event['end'].get('dateTime', event['end'].get('date'))
    date = start_time.split('T')[0]
    color_id = event.get('colorId', 'N/A')
    description = event.get('description', 'N/A')

    # Check if the event is recurring based on the presence of 'originalStartTime'
    is_recurring = 'Yes' if 'originalStartTime' in event else 'No'

    row = [event_id, summary, start_time, end_time, date, color_id, description, is_recurring]
    csv_writer.writerow(row)

def main():
    service = create_calendar_service()
    export_calendar_events(service, calendar_id, csv_file)

if __name__ == '__main__':
    main()
