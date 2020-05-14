"""
This program will read in an excel file produced by my workplace with my schedule
It will parse out the times and dates related to me and store them in objects
It will then create calendar events using Google Calendar API and push them onto
my Google Calendar
"""

import xlrd
from tkinter import filedialog
import datetime
from addCalEvent import create_event

class CalendarEvent:
    """ Object that stores all neccessary information to create a Google Calendar Event

    Attributes:
        date
        start_date_time
        end_date_time
        summary: a.k.a Title
        description
        location
    """

    def __init__(self, date, full_shift, summary='Work', description='Scheduled work hours', location='Carlow'):
        self.date = date
        self.start_date_time = datetime.datetime(date.year, date.month, date.day, int(full_shift[:2]), int(full_shift[2:4]))
        self.end_date_time = self.start_date_time + datetime.timedelta(hours=8)
        self.summary = summary
        self.description = description
        self.location = location

    def __repr__(self):
        return f"""
        Date: {self.date}
        Start: {self.start_date_time}
        End: {self.end_date_time}
        """


def get_users_scheduled_days(worksheet):
    """
    Return a list of all the selected users work times for a month.

    Args:
        worksheet: An xlrd.sheet.Sheet Instance

    Returns:
        A list of working hours for each day in the month
    """
    selected_worksheet = worksheet
    employees_name = 'Aaron'     #input("Who's schedule would you like to generate?: ")
    
    week_of_work_time_cells = []
    for index, cell in enumerate(selected_worksheet.col(0)):
        if cell.value == employees_name:
            week_of_work_time_cells.append(selected_worksheet.row(index))
 
    week_of_work_times = []
    for row in week_of_work_time_cells:
        for index in range(1, len(row) - 2):
            week_of_work_times.append(row[index].value)
    
    return week_of_work_times

def get_worksheet_from_workbook():
    """
    Return an xlrd.sheet.Sheet object by taking input from the user
    """
    # user_selected_workbook = xlrd.open_workbook(filedialog.askopenfilename())
    user_selected_workbook = xlrd.open_workbook('mansch.xlsx')
    for sheet in user_selected_workbook.sheet_names():
        print(sheet)
    user_selected_sheet = 'March 20'   # input('Which sheet would you like to populate from?: ')
    return user_selected_workbook.sheet_by_name(user_selected_sheet)

def get_starting_date(worksheet):
    """ Get the starting date from the top of the worksheet and convert it to datetime

    On our schedule, it provides the week ending date,
    so we must grab this and bring it back 7 days

    Args:
        worksheet: An xlrd.sheet.Sheet Instance

    Returns:
        A datetime.date object
    """

    week_ending_date = worksheet.cell(0, 0).value[-8:]
    year = '20' + week_ending_date[-2:]
    month = week_ending_date[3:5]
    day = week_ending_date[:2]
    week_ending_date_to_obj = datetime.date(int(year), int(month), int(day))
    starting_date = week_ending_date_to_obj - datetime.timedelta(days=7)
    return starting_date

selected_worksheet = get_worksheet_from_workbook()
work_hours_monthly = get_users_scheduled_days(selected_worksheet)
current_running_date = get_starting_date(selected_worksheet)

calendar_events = []
for work_hours in work_hours_monthly:
    if not work_hours.isalnum():
        event = CalendarEvent(current_running_date, work_hours)
        calendar_events.append(event)
    current_running_date += datetime.timedelta(days=1)

for event in calendar_events:
    create_event(event)