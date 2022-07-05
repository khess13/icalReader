from icalendar import Calendar
from datetime import date, datetime, timedelta
import os

root = os.getcwd()+'\\'
today = date.today()

def get_ical_from_folder(filepath, ext='.ics') -> str:
    # get ical from root directory
    files_in_dir = os.listdir(filepath)
    loc_list = [f for f in files_in_dir if ext in f]
    if len(loc_list) > 1 or len(loc_list) == 0:
        return 'Error'
    else:
        return root+loc_list[0]

def retrieve_ical(location) -> Calendar():
    #open ical doc at location
    with open(location,'rb') as f:
        cal_object = Calendar.from_ical(f.read())
    return cal_object

file = get_ical_from_folder(root)
calendar = retrieve_ical(file)


for component in calendar.walk():
    if component.name == 'VEVENT' and component.get('X-MICROSOFT-CDO-BUSYSTATUS') == 'OOF':
        summary = component.get('summary')
        busy_status =  component.get('X-MICROSOFT-CDO-BUSYSTATUS')
        # date and time file was pulled?
        #datestamp = component.get('dtstamp').dt
        start_time = component.get('dtstart').dt
        end_time = component.get('dtend').dt

        print(summary)
        print(busy_status)
        #print(datestamp)
        print(start_time)
        print(end_time)

       #if start_time >= today and end_time <= today:
        # TODO -- compute hours from start/end
        days = end_time - start_time
        print(days.days)
        print(days.hour)
        # TODO -- get for week? just today?
        # API access spreadsheet https://docs.microsoft.com/en-us/graph/excel-write-to-workbook
