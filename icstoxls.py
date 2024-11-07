from config import students, rows
import csv
from datetime import datetime
import openpyxl
import pytz
import locale

# Set the maximum date to include events up to
max_date = datetime(2024, 11, 15).date()

# Set the locale to French
locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

# Create a new workbook and select the active sheet
wb = openpyxl.Workbook()
ws = wb.active

# Write the header row
ws.append(["Name", "Email", "Events"])

# Loop through each student
for student in students:
    # Create an empty dictionary to store the events by day
    events_by_day = {}

    # Open the CSV file
    with open('planning.csv') as csvfile:
        # Create a CSV reader object
        reader = csv.DictReader(csvfile, delimiter=';')

        # Loop through each row in the CSV file
        for row in reader:
            # Add the event to the dictionary if it's for the current student
            if row[rows["email"]] == student:
                event_date = datetime.strptime(row["Début.Événement"], "%d/%m/%Y %H:%M").date()
                if event_date <= max_date:
                    if event_date not in events_by_day:
                        events_by_day[event_date] = []
                    events_by_day[event_date].append(row)

    # Format the events information as a single string for each cell
    events_str = ""
    for date, events in events_by_day.items():
        date_str = date.strftime('%A %d %B')
        events_str += f"\n{date_str}:\n"
        for event in events:
            # Convert the start and end times to Europe/Paris
            start_time = datetime.strptime(event["Début.Événement"], "%d/%m/%Y %H:%M").astimezone(pytz.timezone("Europe/Paris")).strftime('%H:%M')
            end_time = datetime.strptime(event["Fin.Événement"], "%d/%m/%Y %H:%M").astimezone(pytz.timezone("Europe/Paris")).strftime('%H:%M')
            events_str += f"- {event[rows['coursename']]} ({start_time} - {end_time})"
            if event["Salles"]:
                events_str += f" dans la salle {event['Salles']}"
            # Extract the teacher from the description
            teacher = event[rows["teacher"]]
            if teacher != ' ':
                events_str += f" avec {teacher}"
            events_str += "\n"

    # Write the student's name, email, and events to the worksheet
    ws.append([student.split('@')[0], student, events_str])

# Save the workbook to a file
wb.save("student_events.xlsx")
