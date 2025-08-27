import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

CREDENTIALS_FILE = '../credentials.json'

PICKLE_FILE = '../token.pickle'

def get_calendar_service():
   creds = None

   # The file token.pickle stores the user's access and refresh tokens
   if os.path.exists(PICKLE_FILE):
       with open(PICKLE_FILE, 'rb') as token:
           creds = pickle.load(token)

   # If there are no (valid) credentials available, let the user log in.
   if not creds or not creds.valid:
       if creds and creds.expired and creds.refresh_token:
           creds.refresh(Request())
       else:
           flow = InstalledAppFlow.from_client_secrets_file(
               CREDENTIALS_FILE, SCOPES)
           creds = flow.run_local_server(port=0)

       with open(PICKLE_FILE, 'wb') as token:
           pickle.dump(creds, token)

   service = build('calendar', 'v3', credentials=creds)
   return service