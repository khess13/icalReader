import os
import urllib
from icalendar import Calendar
from datetime import date, datetime, timedelta
from dotenv import load_dotenv

load_dotenv()
root = os.getcwd()+'\\'
today = date.today()
calendar_files = root + '\\calendar_files'
calendar_url = os.getenv('TIMESHEET')

# set up calendar folder
if not os.path.isdir(calendar_files):
    os.mkdir(calendar_files)

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
    # TODO leaving everything in memory, may need to change this to files?
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

            #computed metrics
            dayshours = end_time-start_time
            days = dayshours.days
            hours = round(dayshours.total_seconds()/3600,2)

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

        #less than 1 day
        #more than 1 day
        #hours less than 7.5
        #hours more than 7.5
        #create values for days between if more than 1 day fill 23,24,25, etc?

        #if start_time >= today and end_time <= today:
        # TODO -- compute hours from start/end <<<< they write one, but put hrs in comments


        # TODO -- get for week? just today?
        # API access spreadsheet https://docs.microsoft.com/en-us/graph/excel-write-to-workbook
