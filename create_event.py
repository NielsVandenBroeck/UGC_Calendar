from cal_setup import get_calendar_service

def createEvent(startDate, endDate, name, colleagues):
   service = get_calendar_service()
   #mpeei3q77ufnpq04rvl1vrgsm8@group.calendar.google.com   test
   #070k5lmlvnb1ohd7326sab9d8g@group.calendar.google.com   werk Niels
   idDict = {"Roos": "mpeei3q77ufnpq04rvl1vrgsm8@group.calendar.google.com", "Niels": "070k5lmlvnb1ohd7326sab9d8g@group.calendar.google.com"}
   if name in idDict:
       id = idDict[name]
   else:
       id = "070k5lmlvnb1ohd7326sab9d8g@group.calendar.google.com"
   service.events().insert(calendarId=id,
       body={
           "summary": 'UGC',
           "description": colleagues,
           "start": {"dateTime": startDate.isoformat(), "timeZone": 'GMT+2'},
           "end": {"dateTime": endDate.isoformat(), "timeZone": 'GMT+2'},
       }
   ).execute()
