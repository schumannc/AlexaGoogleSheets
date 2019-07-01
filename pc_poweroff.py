#!/usr/bin/python3.7

import pickle
import os.path
from pathlib import Path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import time

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1ESvfqGkzDMf-APi4--RLuuTY1cHJq49msVUirG0Egek'
RANGE_NAME = 'Sheet1!B1'
CWD = Path(__file__).parents[0]

def device_is_on(sheet):
    body = {
                'values': [[1]]
            }
    result = sheet.values().update(
        spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME,
        valueInputOption='USER_ENTERED', body=body).execute()
    value = result.get('values', [[1]])[0][0]
    return value

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    toke_file = CWD / 'token.pickle'
    if os.path.exists(toke_file):
        with open(toke_file, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(toke_file, 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    print('Setting up default values to On (1-On, 0-Off)')
    value = device_is_on(sheet)
    print('PC Power status set to: ', value)
    print('='*31)
    print('Alexa Service On-line')
    print('Running pc power control')
    print('Say: Alexa trigger pc power off')
    print('='*31)
    while True:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=RANGE_NAME).execute()
        value = result.get('values', [[1]])[0][0]
        if int(value) == 0:
            value = device_is_on(sheet)
            # Turn computer off
            print('The system will poweroff in 2 sec')
            time.sleep(2)
            os.system('systemctl poweroff')
        else:
            time.sleep(5)


if __name__ == '__main__':
    main()