from config import students, rows
import csv
from ics import Calendar, Event
from datetime import datetime
import pytz



for student in students:
    print(student)

    # Create a new ICS calendar
    cal = Calendar()

    # Create an empty list to store the events
    events = []

    # Open the CSV file
    with open('planning.csv') as csvfile:
        # Create a CSV reader object
        reader = csv.DictReader(csvfile, delimiter=';')

        # Loop through each row in the CSV file
        for row in reader:
            # Add the event to the list
            if row[rows["email"]] == student:
                events.append(row)

    for event_found in events:
        event = Event()

        # Set the event properties based on the data in the dictionary
        if event_found[rows["coursename"]] != '':
            event.name = event_found[rows["coursename"]]
        elif event_found[rows["description"]] != '':
            event.name = event_found[rows["description"]]
        else:
            event.name = "Évènement"

        # Parse the start and end times using the datetime module
        start_time = datetime.strptime(event_found["Début.Événement"], "%d/%m/%Y %H:%M")
        end_time = datetime.strptime(event_found["Fin.Événement"], "%d/%m/%Y %H:%M")

        # Set the timezone of the start and end times to Europe/Paris
        start_time = pytz.timezone("Europe/Paris").localize(start_time)
        end_time = pytz.timezone("Europe/Paris").localize(end_time)

        # Set the start and end times of the event
        event.begin = start_time
        event.end = end_time

        event.description = f"{event_found[rows["room"]]}\n{event_found[rows["teacher"]]}\n{event_found[rows["description"]]} \n\n{event_found[rows["coursename"]]}\n{event_found[rows["coursetype"]]}\n\nFait avec amour par Rémi."
        event.location = event_found["Salles"]

        # Add the event to the calendar
        cal.events.add(event)

    username = student.split('@')[0]

    filename = f"{username}.ics"

    # Write the calendar to a file
    with open(filename, 'w') as f:
        f.writelines(cal)
