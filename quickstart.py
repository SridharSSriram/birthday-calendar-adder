import csv
import json
import pandas as pd
import datetime

from dateutil import parser
from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools


# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'
FALL_SEM_YEAR = '2018'
SPRING_SEM_YEAR = '2019'
ROSTER_FILE = "5SERoster.csv"
BIRTHDAY_COLUMN_NAME = "BIRTH_DATE"
FIRSTNAME_COLUMN_NAME = "FIRST_NAME"
LASTNAME_COLUMN_NAME = "LAST_NAME"

def main():

    birthday_df = pd.read_csv(ROSTER_FILE)

    # credentials loaded in from token.json
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))
    
    for row in birthday_df.iterrows():
        # dictionary to pass beginning time, end time, and  into create_event function t
        birthday_info = {}
        
        # turning birth_date into datetime
        birth_date = row[1][birthday_df.columns.get_loc(BIRTHDAY_COLUMN_NAME)]
        birth_date_list = birth_date.split('/')
        
        if(int(birth_date_list[0]) > 8):
            birth_date_list[-1] = FALL_SEM_YEAR
        else:
            birth_date_list[-1] = SPRING_SEM_YEAR
        
        birthday = "./".join(birth_date_list)
        birthday = birthday.replace(' ','T')
        birthday = parser.parse(birthday)
        
        # -05:00 indicating UTC timezone
        birthday = str(birthday).replace(' ', 'T')+'-05:00'
        
        full_name = row[1][birthday_df.columns.get_loc(FIRSTNAME_COLUMN_NAME)] +" "+ row[1][birthday_df.columns.get_loc(LASTNAME_COLUMN_NAME)]

        birthday_info["name"] = full_name
        birthday_info["begin"] = json.dumps(birthday, default = myconverter)
        birthday_info["end"] = json.dumps(birthday.replace("00:00:00","23:59:00"), default = myconverter) 
        new_birthday = create_event(birthday_info)
        
        event = service.events().insert(calendarId='primary', body=new_birthday).execute()
        print('Event created: %s' % (new_birthday.get('htmlLink')))

""" below function created to generate json serializable datetimes
        * returning the object string instance allows for more easy serializable data types"""
def myconverter(o):
    if isinstance(o, datetime.datetime):
        return o.__str__()

def create_event(birthday_dict):
    event = {
      'summary': birthday_dict["name"] + ' Birthday',
      'start': {
        'dateTime': birthday_dict["begin"].replace("\"",""),
        'timeZone': 'America/New_York',
      },
      'end': {
        'dateTime': birthday_dict["end"].replace("\"",""),
        'timeZone': 'America/New_York',
      },
      'reminders': {
        'useDefault': True,
      },
    }
    return event


if __name__ == '__main__':
    main()