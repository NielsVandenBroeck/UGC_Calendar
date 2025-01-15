from cal_setup import get_calendar_service

def createEvent(startDate, endDate, name, colleagues):
   service = get_calendar_service()
   #mpeei3q77ufnpq04rvl1vrgsm8@group.calendar.google.com   test
   #070k5lmlvnb1ohd7326sab9d8g@group.calendar.google.com   werk Niels
   idDict = {"Niels": "070k5lmlvnb1ohd7326sab9d8g@group.calendar.google.com",
             "All": "dad76d3f224a4c6778b4e2def8f6fd3c3ea9ea3899fda8dca4b5c3b950bc6b79@group.calendar.google.com"}
   id = idDict[name]
   service.events().insert(calendarId=id,
       body={
           "summary": 'UGC',
           "description": colleagues,
           "start": {"dateTime": startDate.isoformat(), "timeZone": 'Europe/Brussels'},
           "end": {"dateTime": endDate.isoformat(), "timeZone": 'Europe/Brussels'},
       }
   ).execute()
