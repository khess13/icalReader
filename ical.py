from icalendar import Calendar
import os

root = os.getcwd()

def get_ical_from_folder(filepath, ext='.ical') -> string:
    # get ical from root directory
    files_in_dir = os.listdir(filepath)
    loc_list = [f for f in files_in_dir if ext in f]
    if len(loc_list) > 1:
        return 'Error'
    else:
        return loc_list[0]

def retrieve_ical(location) -> Calendar():
    #open ical doc at location
    f = open(location,'rb')
    cal_object = Calendar.from_ical(f.read())

file = get_ical_from_folder(root)
calendar = retrieve_ical(file)

for component in calendar.walk():
    if component.name == 'VEVENT' and component.get('X-MICROSOFT-CDO-BUSYSTATUS'):
        print(component.get('summary'))
        print(component.get('X-MICROSOFT-CDO-BUSYSTATUS'))
        print(component.get('dtstamp').dt)
        print(component.get('dtstart').dt)
        print(component.get('dtend').dt)

        # TODO -- compute hours from start/end
        # TODO -- get for week? just today?
        # API access spreadsheet https://docs.microsoft.com/en-us/graph/excel-write-to-workbook
