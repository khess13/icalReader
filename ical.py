import os
import json
# import re
import urllib.request
import pytz
from icalendar import Calendar
from datetime import date, datetime
from dotenv import load_dotenv
# from pandas import DataFrame
from openpyxl import load_workbook

load_dotenv()
utc = pytz.UTC
ROOT = os.getcwd()+'\\'
TODAY = date.today()
# calendar_files = ROOT + '\\calendar_files'
CALENDAR_URL = os.getenv('TIMESHEET')

# set up calendar folder
# if not os.path.isdir(calendar_files):
#    os.mkdir(calendar_files)
def json_loader(jsonfile) -> str:
    # opens files and returns struct
    with open(jsonfile,'rb') as j:
        return json.load(j)

def get_ical_from_folder(filepath, ext='.ics') -> str:
    # get ical from ROOT directory
    files_in_dir = os.listdir(filepath)
    loc_list = [f for f in files_in_dir if ext in f]
    if len(loc_list) > 1 or len(loc_list) == 0:
        return 'Error'
    else:
        return ROOT+loc_list[0]

def get_ical_from_url(file='calendar_list.json') -> dict:
    # retrieve ical from web make into a dictionary
    # TODO leaving everything in memory, may need to change this to files?
    calendars = {}
    cal = json_loader(file)
    for emp, link in cal.items():
        c = urllib.request.urlopen(link)
        # return Calendar Object in dict
        calendars[emp] = Calendar.from_ical(c.read())
    return calendars

def retrieve_ical_from_file(location) -> Calendar():
    # open ical doc at location
    with open(location, 'rb') as f:
        cal_object = Calendar.from_ical(f.read())
    return cal_object

def return_updated_dayshrs(days, hours) -> str:
    # decision points
    # TODO -- what is this output going to be for?
    # less than 1 day
    if hours <= 24:
        return str(hours)
    # hours less than 7.5
    elif hours <= 7.5:
        return str(hours)
    # hours more than 7.5
    elif hours > 8:
        return str(days)
    # TODO -- create values for days between if more than 1 day fill 23,24,25...


def abs_type(text_from_calendar) -> str:
    # TODO takes string from summary tries to determine what it is
    pass

def get_datetime(component) -> datetime:
    # return dt from component
    # date and time file was pulled?
    # datestamp = component.get('dtstamp').dt
    dstart_time = component.get('dtstart').dt
    dend_time = component.get('dtend').dt
    return dstart_time, dend_time

def get_string_info(component) -> str:
    # return text from component
    dsummary = component.get('summary')
    dbusy_status = component.get('X-MICROSOFT-CDO-BUSYSTATUS')
    return dsummary, dbusy_status

def calendar_filter(component, target_date) -> bool:
    # returns data for further processing
    if component.name == 'VEVENT' and component.get('X-MICROSOFT-CDO-BUSYSTATUS') == 'OOF':
        st_t, en_t = get_datetime(component)
        if st_t.replace(tzinfo=utc) <= target_date <= en_t.replace(tzinfo=utc):
            return True
    return False

# retrieve from file for testing
# file = get_ical_from_folder(ROOT)
# calendar = retrieve_ical_from_file(file)
# retrieve from url
ABS_REASONS = json_loader('abs.json')

# TODO -- steps
cal_dict = get_ical_from_url()
# testing date
DATE = datetime(2022, 7, 1).replace(tzinfo=utc)
# DATE = date.today()

for emp_name, cal in cal_dict.items():
    print(emp_name)
    # calendar = retrieve_ical_from_file(cal)
    calendar = cal
    for component in calendar.walk():
        if calendar_filter(component, DATE):
            start_time, end_time = get_datetime(component)
            summary, busy_status = get_string_info(component)

            # computed metrics
            dayshours = end_time-start_time
            days = dayshours.days
            hours = round(dayshours.total_seconds()/3600, 2)

            print(summary)
            print(busy_status)
            # print(datestamp)
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
'''


# TODO -- get for week? just today?
# API access spreadsheet
# https://docs.microsoft.com/en-us/graph/excel-write-to-workbook --this looks too hard

# accessing workbook files

wb = load_workbook(filename=hc_location)
# select active tab
active = wb.active
# get names with row value
# column mapping for different reasons
# set values for cells
# active['B15'].value = 1
# active['I15'] = '' # clear value from this column if other is set
