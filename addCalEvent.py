from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import pprint
import datefinder

# What the program can access within Calendar
# See more at https://developers.google.com/calendar/auth
scopes = ["https://www.googleapis.com/auth/calendar"]

flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", scopes=scopes)

# Use this to pull the users credentials into a pickle file
#credentials = flow.run_console()
#pickle.dump(credentials, open("token.pkl", "wb"))

# Read the credentials from a saved pickle file
credentials = pickle.load(open("token.pkl", "rb"))

# Build the calendar resource
service = build("calendar", "v3", credentials=credentials)

# Store a list of Calendars on the account
result = service.calendarList().list().execute()
calendar_id = result["items"][0]["id"]

result = service.events().list(calendarId=calendar_id).execute()

def create_event(my_event):
    """
    Create a Google Calendar Event

    Args:
        my_event: CalendarEvent object
    """
    print("Created Event for " + str(my_event.date))

    event = {
        "summary": my_event.summary,
        "location": my_event.location,
        "description": my_event.description,
        "start": {
            "dateTime": my_event.start_date_time.strftime('%Y-%m-%dT%H:%M:%S'),
            "timeZone": "Europe/London",
        },
        "end": {
            "dateTime": my_event.end_date_time.strftime('%Y-%m-%dT%H:%M:%S'),
            "timeZone": "Europe/London",
        },
        "reminders": {
            "useDefault": False,
        },
    }

    return service.events().insert(calendarId=calendar_id, body=event, sendNotifications=True).execute()
   

