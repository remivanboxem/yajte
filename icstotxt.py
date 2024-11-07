import pytz
from ics import Calendar
from datetime import date, timedelta
import pyperclip
import locale

# Set the locale to French
locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

# Set the timezone to Europe/Paris
tz = pytz.timezone('Europe/Paris')

# Load the ICS file
with open('person.ics', 'r') as f:
    cal = Calendar(f.read())

# Get the start and end dates of the calendar
start_date = min(event.begin.date() for event in cal.events)
end_date = max(event.end.date() for event in cal.events)

# Create a dictionary to store the events for each day
events_by_day = {}

# Loop through each day in the calendar
current_date = start_date
while current_date <= end_date:
    # Get the events for the current day
    events = [event for event in cal.events if event.begin.date() <= current_date and event.end.date() >= current_date]

    # Only add the events to the dictionary if there are any
    if events:
        events_by_day[current_date] = events

    # Move to the next day
    current_date += timedelta(days=1)

# Create a string to store the output
output = ""

# Print the events for each day and add them to the output string
for date, events in events_by_day.items():
    date_str = date.strftime('%A %d %B')
    output += f"\n{date_str}:\n"
    for event in events:
        # Convert the start and end times to Europe/Paris
        start_time = event.begin.astimezone(tz).strftime('%H:%M')
        end_time = event.end.astimezone(tz).strftime('%H:%M')
        output += f"- {event.name} ({start_time} - {end_time})"
        if event.location:
            output += f" dans la salle {event.location}"
        # Extract the teacher from the description
        teacher = event.description.split('\n')[1]
        if teacher != ' ':
            output += f" avec {teacher}"
        output += "\n"

# Copy the output string to the clipboard
pyperclip.copy(output)
