import warnings

from googleapiclient.errors import HttpError

from cal_setup import get_calendar_service

service = get_calendar_service()


# def getAllCalendarIds():
#     try:
#         calendar_ids = []
#         page_token = None
#         while True:
#             calendar_list = service.calendarList().list(pageToken=page_token).execute()
#             for calendar_entry in calendar_list.get('items', []):
#                 calendar_ids.append(calendar_entry['id'])
#             page_token = calendar_list.get('nextPageToken')
#             if not page_token:
#                 break
#         return calendar_ids
#     except HttpError as error:
#         warnings.warn(f"Error when retrieving calendar list: {error}")
#         return []


def createEvent(startDate, endDate, cal_id, description):
    try:
       service.events().insert(calendarId=cal_id,
           body={
               "summary": 'UGC',
               "description": description,
               "start": {"dateTime": startDate.isoformat(), "timeZone": 'Europe/Brussels'},
               "end": {"dateTime": endDate.isoformat(), "timeZone": 'Europe/Brussels'},
           }
       ).execute()
    except HttpError as error:
        warnings.warn(f' Error when creating event: {error}')
        raise

def createCalendar(name):
    try:
        calendar = {
            'summary': name,
            'timeZone': 'Europe/Brussels'
        }

        created_calendar = service.calendars().insert(body=calendar).execute()
        return created_calendar['id']
    except HttpError as error:
        warnings.warn(f' Error when creating calendar: {error}')
        raise

def deleteCalendar(cal_id):
    try:
        service.calendars().delete(calendarId=cal_id).execute()
    except HttpError as error:
        warnings.warn(f"Error when deleting calendar {cal_id}: {error}")
        raise
