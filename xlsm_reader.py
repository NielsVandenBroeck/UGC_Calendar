# import openpyxl module
import openpyxl
import datetime
from create_event import createEvent
import warnings

first_name = 'Niels'
last_name = ''
path = '/mnt/c/Users/niels/PycharmProjects/UGC_Calendar/Uurroosters/UURROOSTER JULI 24.xlsm'
wb_obj = openpyxl.load_workbook(path)

def getDatesOfPerson():
    for sheet in wb_obj.worksheets:
        #search rows of Person
        personRow = None
        for row in range(1,sheet.max_row+1):
            string1 = sheet.cell(row=row,column=1).value
            string2 = sheet.cell(row=row,column=2).value
            if string1 is None or string2 is None:
                continue
            if first_name in string1 and last_name in string2:
                print(first_name+' '+last_name+' found in sheet '+str(sheet)+' at row '+ str(row))
                personRow = row
                break
        if personRow is None:
            warnings.warn(first_name+' not found in '+str(sheet))
            continue
        #Get columns with associated dates
        dates = {}
        for col in range(1, sheet.max_column+1):
            if(type(sheet.cell(row=2, column=col).value) is datetime.datetime):
                dates[col] = sheet.cell(row=2, column=col).value.replace(year=2024)



        #match columns with row of person and get exact times of events
        for col in dates:
            startTime = sheet.cell(row=personRow, column=col).value
            endTime = sheet.cell(row=personRow+1, column=col).value
            if startTime is None:
                continue
            print(str(dates[col])+': '+str(startTime)+'-'+str(endTime))
            startDate = datetime.datetime.combine(dates[col],startTime)
            endDate = datetime.datetime.combine(dates[col],endTime)
            colleagues = ''
            for colleague in range(1,sheet.max_row+1):
                name = sheet.cell(row=colleague, column=1).value
                if name is None or sheet.cell(row=colleague, column=2).value is None:
                    continue
                startTimeColleague = sheet.cell(row=colleague, column=col).value
                endTimeColleague = sheet.cell(row=colleague+1, column=col).value
                if startTimeColleague is None:
                    continue
                if str(endTimeColleague) == '18:00:00' and str(startTime) == '18:30:00':
                    continue
                if str(endTime) == '18:00:00' and str(startTimeColleague) == '18:30:00':
                    continue
                if sheet.cell(row=colleague+5, column=col).value == "vip":
                    name += ' (vip)'
                colleagues += name + ', '

            createEvent(startDate,endDate, first_name, colleagues)

if __name__ == '__main__':
   getDatesOfPerson()