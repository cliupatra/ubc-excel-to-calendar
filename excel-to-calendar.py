import openpyxl as xl
from openpyxl import Workbook

wb = xl.load_workbook("View_My_Courses.xlsx")
sheet = wb['View My Courses']
new_schedule = Workbook()
schedule_sheet = new_schedule.active
days_of_week = ["Mon", "Tue", "Wed", "Thu", "Fri"]
term1 = {}
term2 = {}

# creates day headings for calendar
def create_headings():
    schedule_sheet.cell(1, 1).value = "Monday"
    schedule_sheet.cell(1, 2).value = "Tuesday"
    schedule_sheet.cell(1, 3).value = "Wednesday"
    schedule_sheet.cell(1, 4).value = "Thursday"
    schedule_sheet.cell(1, 5).value = "Friday"


# examines meeting day cell on excel sheets and extracts the meeting days as a list
def find_class_days(meeting_day_value):
    if meeting_day_value is None:
        return None
    else:
        days = []
        for day in range(5):
            if days_of_week[day] in meeting_day_value:
                days.append(days_of_week[day])
    return days


# returns corresponding column number based on day
def return_day_column(day):
    if day == "Mon":
        return 1
    elif day == "Tue":
        return 2
    elif day == "Wed":
        return 3
    elif day == "Thu":
        return 4
    elif day == "Fri":
        return 5


def find_empty_row(column):
    for row in range(1, schedule_sheet.max_row + 1):
        if schedule_sheet.cell(row, column).value == None:
            return row
    return schedule_sheet.max_row + 1


def place_course_in_calendar(class_name, days):
    if days is not None:
        for day in days:
            column = return_day_column(day)
            row = find_empty_row(column)
            schedule_sheet.cell(row, column).value = class_name


def find_term(meeting_day_value):
    if meeting_day_value is not None:
        if "2024-09" in meeting_day_value:
            return 1
        elif "2025-01" in meeting_day_value:
            return 2

def create_term_dicts():
    for row in range(4, sheet.max_row + 1):
        class_name = sheet.cell(row, 2).value
        days = find_class_days(sheet.cell(row, 8).value)
        term = find_term(sheet.cell(row, 8).value)
        if term == 1:
            term1[class_name] = days
        elif term == 2:
            term2[class_name] = days

create_headings()
create_term_dicts()

for school_class in term1:
    place_course_in_calendar(school_class, term1.get(school_class))



new_schedule.save("new-schedule.xlsx")
