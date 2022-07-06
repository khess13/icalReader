from icalendar import Calendar
from datetime import date, datetime, timedelta
import os
import urllib

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

def get_ical_from_url(file = 'calendar_list.json') -> dict:
    # retrieve ical from web make into a dictionary
    calendars = {}
    with open(file,'r') as j:
        cal = json.load(j)
    for emp, link in cal.items():
        c = urllib.urlopen(link)
        calendars[emp] = c.read()
    return calendars

def retrieve_ical_from_file(location) -> Calendar():
    #open ical doc at location
    with open(location,'rb') as f:
        cal_object = Calendar.from_ical(f.read())
    return cal_object


# retrieve from file for testing
file = get_ical_from_folder(root)
calendar = retrieve_ical_from_file(file)
# retrieve from url
# TODO -- steps
# cal_dict = get_ical_from_url()

# TODO - cycle through several calendars
'''
for emp_name, calendar in cal_dict.items():
    print(emp_name)
    for component in calendar.walk():
        if component.name == 'VEVENT' and component.get('X-MICROSOFT-CDO-BUSYSTATUS') == 'OOF':
            summary = component.get('summary')
            busy_status =  component.get('X-MICROSOFT-CDO-BUSYSTATUS')
            # date and time file was pulled?
            # datestamp = component.get('dtstamp').dt
            start_time = component.get('dtstart').dt
            end_time = component.get('dtend').dt

            print(summary)
            print(busy_status)
            #print(datestamp)
            print(start_time)
            print(end_time)
'''

# Testing 1 loop
for component in calendar.walk():
    if component.name == 'VEVENT' and component.get('X-MICROSOFT-CDO-BUSYSTATUS') == 'OOF':
        summary = component.get('summary')
        busy_status =  component.get('X-MICROSOFT-CDO-BUSYSTATUS')
        # date and time file was pulled?
        # datestamp = component.get('dtstamp').dt
        start_time = component.get('dtstart').dt
        end_time = component.get('dtend').dt
        #debug output
        print(summary)
        #print(busy_status)
        #print(datestamp)
        print(start_time)
        print(end_time)

        # calculated metrics
        dayshours = end_time-start_time
        days = dayshours.days
        hours = round(dayshours.total_seconds()/3600,2)
        print(days)
        print(hours)

       #if start_time >= today and end_time <= today:
        # TODO -- compute hours from start/end <<<< they write one, but put hrs in comments


        # TODO -- get for week? just today?
        # API access spreadsheet https://docs.microsoft.com/en-us/graph/excel-write-to-workbook
