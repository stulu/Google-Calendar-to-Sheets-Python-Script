# Authorisation with no external call.
import datetime
import gspread
import schedule
import json
import time
import os
import tkinter as tk

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tkinter import messagebox

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar',
          'https://www.googleapis.com/auth/spreadsheets']


def main():
    """Shows basic usage of the Google Calendar API.
    Prints the name of the events happening today to a Google Sheet.
    """
    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except google.auth.exceptions.RefreshError:
                tk.messagebox.showerror(
                    "Error", "Your authorization token has expired or been revoked. Please re-authorize the application.")
                return
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API- Argentine time is -3 hours from UTC
        now = datetime.datetime.utcnow()
        start_of_day = datetime.datetime(
            now.year, now.month, now.day, 0, 0, 0, 0, datetime.timezone.utc)
        end_of_day = datetime.datetime(
            now.year, now.month, now.day, 23, 59, 59, 999999, datetime.timezone.utc)
        print(f'Getting the events happening today')
        events_result = service.events().list(calendarId='primary', timeMin=start_of_day.isoformat(),
                                              timeMax=end_of_day.isoformat(), singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])

        # Shows a pop up saying there are no events
        if not events:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("No upcoming events",
                                "There are no events happening today.")
            return

        # Create a list of event names
        events_today = [event['summary'] for event in events]

        # Authenticate and open the Google Sheet
        sa = gspread.service_account()
        sh = sa.open("SHEET NAME")
        wks = sh.sheet1

        # Get the data from the Google Sheet
        data = wks.get_all_values()
        not_found_events = []

        # Check if each event name is in the data retrieved from the Google Sheet
        for event in events_today:
            event_found = False
            for row in data:
                if event in row:
                    # print(f"{event} is already in the Google Sheet")
                    count = int(row[2]) + 1
                    wks.update_cell(data.index(row) + 1, 3, count)
                    event_found = True
                    break
            if not event_found:
                print(f"{event} is not in the Google Sheet")
                not_found_events.append(event)

        root = tk.Tk()
        root.withdraw()
        if not_found_events:
            messagebox.showinfo(
                "Error", f"The following events were not found:\n\n{', '.join(not_found_events)}")
        # Show a pop-up message indicating that the script has been run
        root = tk.Tk()
        root.withdraw()
        tk.messagebox.showinfo("Success", "The script has been run!")

    except HttpError as error:
        print('An error occurred: %s' % error)


"""If you want to run the script with VSC running.
    You can execute the script below.
    """
# Schedule the script to run every day at 8:00 PM
# schedule.every().day.at('18:00').do(main)

# Run the scheduled task at startup if it's already past the scheduled time
# if datetime.datetime.now().hour >= 18:
#    main()

# Wait until the scheduled time
# while datetime.datetime.now().hour < 18:
#    time.sleep(60)

# Run the scheduled task once at the scheduled time
# schedule.run_pending()
if __name__ == '__main__':
    main()
