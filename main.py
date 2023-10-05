from win32com.client import Dispatch
from datetime import datetime, timedelta
from notion.client import NotionClient
from datetime import date
import datetime
from notion.collection import NotionDate
from win10toast import ToastNotifier
import secrets

def add_event_to_notion(event):
    """
    Adds a new row to the Notion database with the given information.

    Args:
    rij (list): A list containing the following information in order:
        - Name (str): The name of the meeting.
        - Start time (str): The start time of the meeting in the format "YYYY-MM-DD HH:MM".
        - End time (str): The end time of the meeting in the format "YYYY-MM-DD HH:MM".
        - Timezone (str): The timezone of the meeting.
        - URL (str): The URL of the meeting.
    
    Returns:
    None
    """
    new_row = collection_view.collection.add_row()
    new_row.Name = event[0]
    new_row.When = NotionDate(start=datetime.datetime.strptime(event[1], "%Y-%m-%d %H:%M"),
                              end=datetime.datetime.strptime(
                                  event[2], "%Y-%m-%d %H:%M"),
                              timezone=event[3],
                              reminder=True
                              )
    new_row.Type = 'Daily meeting'
    new_row.Project = 'Project'
    new_row.URL = event[5]
    new_row.Weekday = week_days[(datetime.datetime.strptime(
        event[2], "%Y-%m-%d %H:%M")).weekday()]
    new_row.Addedby = 'Python'


import datetime

def filterdate(time_list):
    """
    Filters a list of dates to only include those within the current day.

    Args:
    time_list (list): A list of datetime objects.

    Returns:
    bool: True if the datetime object is within the current day, False otherwise.
    """
    today_start = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    today_end = datetime.datetime.now().replace(hour=23, minute=59, second=59, microsecond=999999)

    if time_list >= today_start and time_list <= today_end:
        return True
    else:
        return False


toaster = ToastNotifier()
toaster.show_toast("Notion update", "Your script has been started to run")
client = NotionClient(token_v2=secrets.token)
calendar_url = secrets.calendar_urls

collection_view = client.get_collection_view(calendar_url)
week_days = ["Monday", "Tuesday", "Wednesday",
             "Thursday", "Friday", "Saturday", "Sunday"]
outlook_format = '%Y-%m-%d %H:%M'
outlook = Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
prev_value = ["ab", "ba"]
appointment_df = []
start_time = datetime.time(0)
end_time = datetime.time(23)

day = date.today()
start = day - timedelta(days=day.weekday())
end = start + timedelta(days=10)
start_datetime = (datetime.datetime.combine(
    start, start_time)).strftime("%Y-%m-%d %H:%M")
end_datetime = datetime.datetime.combine(
    end, end_time).strftime("%Y-%m-%d %H:%M")


appointments = ns.GetDefaultFolder(9).Items
appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"

# filter to the range: from = (today - 10), to = (today)
appointments = appointments.Restrict(
    "[Start] >= '" + start_datetime + "' AND [End] <= '" + end_datetime + "'")

# Iterate through restricted AppointmentItems and create a df
for appointment_item in appointments:
    if (appointment_item.Start.Format(outlook_format) > start_datetime) & (
            appointment_item.Start.Format(outlook_format) < end_datetime):
        row = []
        row.append(appointment_item.Subject)
        row.append(appointment_item.Start.Format(outlook_format))
        row.append((appointment_item.Start +
                   timedelta(minutes=appointment_item.Duration)).Format(outlook_format))
        row.append('Central European Time (UTC+01:00)')
        row.append({'unit': 'minute', 'value': 30})
        row.append(appointment_item.body)
        appointment_df.append(row)

i = 0
for event in appointment_df:
    add_event_to_notion(event)
    i = i + 1
notification_status = str(i) + " new row(s) added into Notion"
toaster.show_toast("Notion update", notification_status)

for row in appointment_df:
    subject = row[0]
    if len(collection_view.collection.get_rows(search=subject)) > 1:
        for row in collection_view.collection.get_rows(search=subject):
            if row.name == prev_value[0] and NotionDate.to_notion(row.When) == prev_value[1]:
                row.remove()
            prev_value = [row.Name.strip(), NotionDate.to_notion(row.When)]

toaster.show_toast("Ready!", "Your notion is up-to-date!")
