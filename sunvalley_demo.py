"""A demo script to showcase the Sun Valley ttk theme."""

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import sv_ttk
import sv_ttk
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime
import pandas as pd
import json
import glob
import pickle
# general imports
import requests
from bs4 import BeautifulSoup 
import re
from html import unescape
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os
from tabulate import tabulate
import numpy as np
import json
import pickle
import glob
from ctypes import windll


# Imports related to google API
from google.oauth2 import credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request


import gspread
import gspread_dataframe as gd
import base64

# Library for fuzzy comparison of strings
from fuzzywuzzy import fuzz
from fuzzywuzzy import process



with open('config.json', 'r') as file:
    config = json.load(file)
    # Mapping of abbreviated months to full months
    month_mapping = config['month_mapping']
    # Threshold for fuzzy comparison of strings
    fuzzy_threshold = config['fuzzy_threshold']
    # ID of the google drive folder to work in 
    folder_id = config['folder_id']
    # ID of the google apps script to run when a form is created
    script_id = config['script_id']
    
athlete_status = pd.read_pickle('athlete_status.pkl')

def ordinal_suffix(day):
    return str(day)+("th" if 4<=day%100<=20 else {1:"st",2:"nd",3:"rd"}.get(day%10, "th"))

def dtStylish(dt, compact=False):
    return dt.strftime(f"%a the {{th}}{' of %B %Y' if not compact else ''} ").replace("{th}", ordinal_suffix(dt.day))

def load_previous_events_from_file():
    with open(f"{glob.glob('events*.pkl')[0]}", "rb") as file:
        events_list = pickle.load(file)
    return events_list

def save_events_to_file(events_list):
    with open(f"events_{datetime.today().strftime('%Y_%m_%d')}.pkl", "wb") as file:
        pickle.dump(events_list, file)
        
def get_new_responses(do_print=False):
    SKIP_OLD_DATA = True

    # Load the JSON file into a dictionary
    with open('config.json', 'r') as file:
        config = json.load(file)
    # Load variables from dictionary
    form_id = config['form_id']
    last_access = datetime.fromisoformat(config['last_access'])
    sheet_id = config['sheet_id']
    
    try:
        form_responses = forms_service.forms().responses().list(formId=form_id).execute()
    except Exception as e:
        print(e)
            
    try:
        if SKIP_OLD_DATA:
            # Write the current time back
            config['last_access'] = datetime.now().isoformat()
            # Save the updated dictionary back to the JSON file
            with open('config.json', 'w') as file:
                json.dump(config, file) 
        rows = []
        for form_response in form_responses['responses']:
            # createTime is when the form response was submitted. 
            createTime = form_response['createTime']
            createTime_dt = datetime.strptime(createTime, '%Y-%m-%dT%H:%M:%S.%fZ')
            # Skip responses that have already been stored in the database. Adding 2 hrs because google is 2 hrs in the past.
            if createTime_dt + timedelta(hours=2) < last_access:
                if do_print:
                    print("Response already seen.")
                if SKIP_OLD_DATA:
                    continue 
            # Now the answers are extracted
            __, answer_data = zip(*form_response['answers'].items())
            answers = [(answer['questionId'], answer['textAnswers']['answers']) for answer in answer_data]
            # The answers are sorted by their question ids. For some reason, the google API shuffles the answers. 
            # The answers are 'unshuffled' here, pairing question ids with the order of those ids stored earlier.
            question_order = config['question_order']
            question_order = [str(item) for item in question_order.split(" ")]

            answers_dict = {key: value for key, value in answers}
            answers = [answers_dict.get(key, [{'value': 'None'}]) for key in question_order]
            # Now when someone has signed up for more than one competition, that field 
            # contains multiple values, while all other fields only contain one value.
            # The other values are duplicated so the list-with-a-bump becomes a rectangle. 
            for i in range(len(answers[1])-1):
                answers[0].append(answers[0][0])
                answers[2].append(answers[2][0])
                answers[3].append(answers[3][0])
                answers[4].append(answers[4][0])
                answers[5].append(answers[5][0])
                answers[6].append(answers[6][0])
                answers[7].append(answers[7][0])
                answers[8].append(answers[8][0])
                
            # The answers are still hidden in a few layers. They are unpacked here. 
            answers = np.asarray(answers)
                
            for i in range(np.shape(answers)[1]):
                name = answers[0][i].get('value', None)
                competition = answers[1][i].get('value', None)
                weightclass = answers[2][i].get('value', None)
                squat = answers[3][i].get('value', None)
                bench = answers[4][i].get('value', None)
                deadlift = answers[5][i].get('value', None)
                coach_name = answers[6][i].get('value', None)
                coach_number = answers[7][i].get('value', None)
                coach_dob = answers[8][i].get('value', None)
                status = find_name(name)
                if do_print:
                    print(f"status: {status}")
                row = {
                    'Time of Response': createTime,
                    'Name': name,
                    'Competition': competition,
                    'Weightclass': weightclass,
                    'Squat': squat,
                    'Bench': bench,
                    'Deadlift': deadlift,
                    'Coach Name': coach_name, 
                    'Coach Number' : coach_number,
                    'Coach DOB': coach_dob,
                    'Doping check': status[0],
                    'Student Status': status[1]
                }
                rows.append(row)
        # Create a DataFrame with the new responses.
        df = pd.DataFrame(rows)
        if do_print:
            print(tabulate(df, headers='keys', tablefmt='psql'))
            
        # Store the new responses in the google sheet

        
        sh = gc.open_by_key(sheet_id)
        worksheet = sh.worksheets()[0]
        gd.set_with_dataframe(worksheet=worksheet,dataframe=df,include_index=False,include_column_header=False,row=next_available_row(worksheet),resize=False)  
        
    except KeyError as e:
        print(e)
        print("No responses yet.")
        
# This function is used to get the status of all known athletes of the association.
# It opens an email for the most recent competition invitation. This contains a list of athletes and their status. 
def get_athlete_status(do_print=False):
    # The first step is reading the message of the email into a string.
    try:
        results = mail_service.users().messages().list(maxResults= 2, q="[info] Uitnodiging", userId='me').execute()
        messages = results.get('messages',[])
        for message in messages:
            message = mail_service.users().messages().get(userId='me', id=message['id']).execute()
            email_data = message['payload']['headers']
            for values in email_data:
                name = values['name']
                if name == 'From':
                    from_name= values['value']                
                    for part in message['payload']['parts']:
                        try:
                            data = part['body']["data"]
                            byte_code = base64.urlsafe_b64decode(data)
                            text = byte_code.decode("utf-8")                                        
                        except BaseException as error:
                            print(error)
                            pass                      
    except HttpError as error:
        print(f'An error occurred: {error}')
        
    # Now we look for the link of interest in the email. It is coincidentaly the longest link in the email.
    main_string = str(text)
    keyword = "https://knkf-sectiepowerliften.nl"
    ending_symbol = '"'

    pattern = re.escape(keyword) + r"(.*?)" + re.escape(ending_symbol)
    matches = re.findall(pattern, main_string)

    longest_length = 0
    longest_match = ""

    for match in matches:
        extracted_text = match.strip()
        match_length = len(extracted_text)
        
        if match_length > longest_length:
            longest_length = match_length
            longest_match = extracted_text
    if do_print:
        print("Longest match:", keyword+longest_match)
    web_page_url = keyword+longest_match
    # The link in the e-mail is to a launch page. We need to go to the actual page with the members list.
    web_page_url = web_page_url.replace("club.php", "club-members.php")
    response = requests.get(web_page_url)
    html = response.text
    # Now we parse the html.
    soup = BeautifulSoup(html, 'html.parser')
    # Find elements based on class name "post"
    post_elements = soup.find_all('div', class_='panel-body')
    # From manual inspection the athletes are found to be contained in li elements.
    # Regex is used to find all li elements.
    keyword = "<li>"
    ending_symbol = "</li>"
    pattern = re.escape(keyword) + r"(.*?)" + re.escape(ending_symbol)

    extracted_list = []

    # These elements are put into a list.
    for index, post in enumerate(post_elements, start=1):
        matches = re.findall(pattern, get_full_structure(post))
        for match in matches:
            extracted_text = match.strip()
            extracted_list.append(extracted_text)
            
    athletes = []
    rows = []
    # Each element is again parsed with BeautifulSoup to extract the name and status of the athlete.
    for text in extracted_list:
        sub_soup = BeautifulSoup(text, "html.parser")
        name = sub_soup.find(string=True, recursive=False).strip()
        if do_print:
            print(f'Name = {name}')
        try:
            deelnemen_coachen = sub_soup.select('a[href*="/anti-doping/"]')[0].next_sibling.strip()
            deelnemen_coachen = ["deelnemen" in deelnemen_coachen, "coachen" in deelnemen_coachen]
        except IndexError:
            deelnemen_coachen = [False, False]
        try: 
            studentenstatus = sub_soup.select('a[href*="/studentenstatus/"]')
            studentenstatus = len(studentenstatus) > 0 
        except IndexError:
            studentenstatus = [False]
        rows.append({'Name':name.replace(" (Wedstrijdlid)", "").strip(), 'Participate': deelnemen_coachen[0], 'Coaching': deelnemen_coachen[1], 'Student': studentenstatus})
    if do_print:
        for row in rows:
            print(row)
    df = pd.DataFrame(rows)
    df.to_pickle('athlete_status.pkl') 
    if do_print:
        print(tabulate(df, headers='keys', tablefmt='psql'))
        
    # Load the JSON file into a dictionary
    with open('config.json', 'r') as file:
        config = json.load(file)
    # Load variables from dictionary
    sheet_id = config['sheet_id']
    
    sh = gc.open_by_key(sheet_id)
    worksheet = sh.worksheets()[1]
    gd.set_with_dataframe(worksheet=worksheet,dataframe=df,include_index=False,include_column_header=False,row=2,resize=False)

def update_sheets():
    get_new_responses()
    get_athlete_status()

regex_pattern_two_days = r'(\d{1,2})\s*-\s*(\d{1,2})\s*([a-zA-Z]+)'
# This function expects a string of the format "## - ## AAA." where # represents a number from 1-31 and A represents any letter
# I use regex to convert the string to a datetime object
def get_signup_date(date_string):
    try:
        # The regular expression splits the string into the starting date, ending date and abbreviated month
        start_day, end_day, month_abbrev = re.match(regex_pattern_two_days, date_string).groups() 
    except AttributeError:
        # In case the event is only one day, there is only one date given, so we can just split
        #print(f"Found single day: {date_string.split(' ')}")
        start_day, month_abbrev = date_string.split(" ")
        end_day = start_day

    # Get the current year since it's not included in the string
    current_year = datetime.now().year

    # Create datetime objects
    start_date = datetime(current_year, list(month_mapping.keys()).index(month_abbrev) + 1, int(start_day))
    end_date = datetime(current_year, list(month_mapping.keys()).index(month_abbrev) + 1, int(end_day))
    
    # Now if we find that the upcoming event is in the past, that means that the event will actually be next year instead of this year. 
    # So we take the same month and day but just increment the year by 1.
    if datetime.now() > start_date:
        start_date = datetime(current_year+1, list(month_mapping.keys()).index(month_abbrev) + 1, int(start_day))
        end_date = datetime(current_year+1, list(month_mapping.keys()).index(month_abbrev) + 1, int(end_day))
        
    # The KNKF has put on their website that they will open the competition for signup 60 days before the competition. 
    signup_date = start_date-timedelta(days=60)
    return signup_date, start_date, end_date

       
# Function to create or get credentials for google API services.
def get_credentials(type):
    
    if type.lower() == "drive":
        SCOPES = ["https://www.googleapis.com/auth/drive.file"]  # Use the appropriate scope for the Google Drive API
        TOKEN_FILE = "drive_token.json"
    elif type.lower() == "forms":
        SCOPES = ['https://www.googleapis.com/auth/forms.body', 
                  'https://www.googleapis.com/auth/forms.responses.readonly']
        TOKEN_FILE = "form_token.json"
    elif type.lower() == "scripts":
        SCOPES = ['https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/forms']
        TOKEN_FILE = "scripts_token.json"
    elif type.lower() == "read_mail":
        SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify', 'https://mail.google.com/']
        TOKEN_FILE = "read_mail_token.json"
    else:
        print("Unsupported credential type. Please enter a supported type or implement the missing type.")
        return None
    creds = None
    # If valid credentials already exist, just take them from the file    
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    # Else, create or refresh them, and store them in the file
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("client_secrets.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
    return creds

def find_name(name):
    # Apply the function to the DataFrame and get the best match
    athlete_status["SimilarityScore"] = athlete_status["Name"].apply(fuzz.token_sort_ratio, args = (name,))
    best_match_row = athlete_status[athlete_status["SimilarityScore"] >= fuzzy_threshold].sort_values(by="SimilarityScore", ascending=False).iloc[0]

    # Extract the best match and its similarity score
    best_match = best_match_row["Name"]
    similarity_score = best_match_row["SimilarityScore"]

    # Compare the similarity score with the threshold
    if similarity_score >= fuzzy_threshold:
        print(f"Closest match: {best_match} (Similarity: {similarity_score}%)")
        status = athlete_status.loc[athlete_status["Name"]== best_match]
        return status[["Participate", "Student"]].values.tolist()[0]
    else:
        print("No acceptable match found.")
        return ["ERROR", "Name not found"]
    
# Function to find the next empty row in a google worksheet
def next_available_row(worksheet):
    return len(worksheet.col_values(1))+1

# Function to recursively get the entire HTML structure of an element
def get_full_structure(element):
    return "".join([str(tag) for tag in element.contents])



class Event:
    def __init__(self, name, signup_date, start_date, end_date, location, organisation, link, notes=None):
        self.name = name
        self.signup_date = signup_date
        self.start_date = start_date
        self.end_date = end_date
        self.location = location
        self.organisation = organisation
        self.link = link
        self.notes = notes
        
    def __str__(self):
        return (
            f"Event Name: {self.name}\n"
            f"Early Signup Deadline: {dtStylish(self.signup_date)}\n"
            f"Date of Event: {dtStylish(self.start_date, True)}to {dtStylish(self.end_date)}\n"
            f"Location: {self.location}\n"
            f"Organisation: {self.organisation}\n"
            f"Link: {self.link}\n"
            f"Notes: {self.notes}\n"
        )
    
    def __eq__(self, other):
        if isinstance(other, Event):
            return (
                self.name == other.name and
                self.signup_date == other.signup_date and
                self.start_date == other.start_date and
                self.end_date == other.end_date and
                self.location == other.location and
                self.organisation == other.organisation and
                self.link == other.link and
                self.notes == other.notes
            )
        return False
            
    def event_to_dataframe(self):
        data = {
            "Name": [self.name],
            "Signup Date": [self.signup_date],
            "Start Date": [self.start_date],
            "End Date": [self.end_date],
            "Location": [self.location],
            "Organisation": [self.organisation],
            "Link": [self.link],
            "Notes": [self.notes]
        }
        df = pd.DataFrame(data)
        return df
    
    @classmethod
    def from_dataframe(cls, df):
        return cls(
            name=df["Name"].iloc[0],
            signup_date=df["Signup Date"].iloc[0],
            start_date=df["Start Date"].iloc[0],
            end_date=df["End Date"].iloc[0],
            location=df["Location"].iloc[0],
            organisation=df["Organisation"].iloc[0],
            link=df["Link"].iloc[0],
            notes=df["Notes"].iloc[0]
        )
        
        
# Setup authentication and authorization with the drive and forms APIs
drive_creds = get_credentials("drive")
forms_creds = get_credentials("forms")
mail_creds = get_credentials("read_mail")
scripts_creds = get_credentials("scripts")

DRIVE_DISCOVERY_DOC = "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"
FORMS_DISCOVERY_DOC = "https://forms.googleapis.com/$discovery/rest?version=v1"
MAIL_DISCOVERY_DOC = "https://gmail.googleapis.com/$discovery/rest?version=v1" 

drive_service = build("drive", "v3", credentials=drive_creds, discoveryServiceUrl=DRIVE_DISCOVERY_DOC)
forms_service = build('forms', 'v1', credentials=forms_creds, discoveryServiceUrl=FORMS_DISCOVERY_DOC, static_discovery=False)
mail_service = build('gmail', 'v1', credentials=mail_creds, discoveryServiceUrl=MAIL_DISCOVERY_DOC)
scripts_service = build('script', 'v1', credentials=scripts_creds) # for some reason this service still works without a discovery doc



gc = gspread.oauth(
    credentials_filename="client_secrets.json",
    authorized_user_filename='gspread_token.json'
)     

def write_wa_msg(events):
    # Load the JSON file into a dictionary
    with open('config.json', 'r') as file:
        config = json.load(file)
    form_id = config['form_id']
    # Define input and output file paths
    input_file_path = 'whatsapp_msg_def.txt'
    output_file_path = 'whatsapp_msg.txt'

    # Concatenate string representations of objects
    events_string = '\n'.join(str(event) for event in events)
    form_id_string = f"https://docs.google.com/forms/d/{form_id}/viewform"

    # Read the default content
    with open(input_file_path, 'r') as input_file:
        default_content = input_file.read()

    # Replace the first occurrence of {} with replacement_string1
    modified_content1 = default_content.replace('{}', events_string, 1)

    # Replace the second occurrence of {} with replacement_string2
    modified_content2 = modified_content1.replace('{}', form_id_string, 1)

    # Write the modified content to the output file
    with open(output_file_path, 'w') as output_file:
        output_file.write(modified_content2)



def find_events():
    # Fetch the web page
    web_page_url = 'https://knkf-sectiepowerliften.nl/kalender/'
    response = requests.get(web_page_url)
    html = response.text
    # Parse the HTML
    soup = BeautifulSoup(html, 'html.parser')

    # Find elements based on class name "post"
    post_elements = soup.find_all('section', class_='post')

    # List to store event objects
    events = []
    # We have two kinds of posts on the page. They vary enough that they should be handled separately.
    # The first type is for competitions coming soon; they are already open for signup with the KNKF.
    for post in post_elements:
        if not "style" in post.attrs:
            continue    # These posts have a style attribute in the html, while the later competitions do not
        
        name_element = post.find('h2').find('a')
        event_name = name_element.text.strip() if name_element else ""
        if "Masters" in event_name:
            continue    # Since we never have Masters athletes, we don't need to announce these competitions
        
        event_link = name_element['href'] if name_element else ""
        
        date_location_element = post.find('p')
        event_date_location = date_location_element.get_text(strip=True) if date_location_element else ""
        event_date, event_location = event_date_location.split("•") if "•" in event_date_location else ("", "")
        signup_date, start_date, end_date = get_signup_date(event_date)
        # event_date = event_date.split(" ")
        # event_date = month_mapping[event_date[3][:-1]] + " " + "".join(event_date[0:3]) + " " + event_date[4] 
        
        organisation_element = post.find('p', class_='intro')
        event_organisation = organisation_element.text.strip() if organisation_element else "No organisation found"
        # Create Event object and add it to the list
        event = Event(event_name.strip(), signup_date, start_date, end_date, event_location.strip(), event_organisation.strip()[13:], event_link.strip(), "Officially open for signup with the KNKF. Don't wait too long!")
        events.append(event)
        
    # Extract information from the competitions coming later. These are defined in the last post
    post = post_elements[-1]
    date_pattern = re.compile(r'\d{1,2}\s+[a-zA-Z&]{2,}')  # Date pattern in Dutch
    date_paragraphs = []

    for paragraph in post.find_all('p'):
        #print(paragraph)
        paragraph_text = unescape(paragraph.get_text())  # Decode HTML entities
        #print(paragraph_text)
        if "Wedstrijdinschrijvingen openen 60 dagen" in paragraph_text:
            break
        if date_pattern.search(paragraph_text):
            date_paragraphs.append(paragraph_text.split('\n'))

    for paragraph in date_paragraphs:
        event_notes = []
        #print(f"par split returns: {paragraph[0].split(':')}")
        event_date, event_name = paragraph[0].split(':')
        if any([word in event_name.lower() for word in ["masters", "vergadering"]]):
            continue
        event_date = event_date.replace("&", "-")
        if "-" in event_date:
            event_date = event_date.split()
            event_date = " ".join(event_date[0:-1]) + " " + event_date[-1][0:3]
        else:
            event_date = event_date.split(" ")
            event_date = event_date[0] + " " + event_date[1][0:3]
        #print(event_date)
        signup_date, start_date, end_date = get_signup_date(event_date)
        if signup_date < events[-1].signup_date:
            signup_date += timedelta(days=365)
            start_date += timedelta(days=365)
            end_date += timedelta(days=365)
        if '&' in event_name:
            event_name = event_name.split('&')[0]
        location_line = paragraph[1].split('-')
        if len(location_line) > 2:
            event_notes.append(location_line[2].strip().capitalize() +".")
        event_location = location_line[0]
        event_organisation = location_line[1]
        if len(paragraph) > 2:
            event_notes.append(paragraph[2].strip().capitalize())
        event_notes = ' '.join(event_notes)
        if event_notes == "":
            event_notes = "None."
        if start_date > datetime.now() + relativedelta(months=6):
            break
        event = Event(event_name.strip(), signup_date, start_date, end_date, event_location.strip(), event_organisation.strip(), "No link available yet", event_notes)
        events.append(event)
    return events

def create_new_form(events):
    # Form 'updates': all modifications to go from the default form to the required form 

    # Request body for creating a Google Form file in the specified folder
    NEW_FORM = {
        "name": "Competition Signup Form",
        "parents": [folder_id],
        "mimeType": "application/vnd.google-apps.form"
    }

    with open("form_description.txt", "r", encoding="utf-8") as file:
        new_description = file.read()

    update_description = {
        "requests" : [{
            "updateFormInfo" : {
                "info" : {
                    "description": new_description
                },
                "updateMask" : "description"
            }
        }]
    }

    name_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "What is your name?",
                    "questionItem": {
                        "question": {
                            "required": True,
                            "textQuestion": {}
                        }
                    },
                },
                "location": {
                    "index": 0
                }
            }
        }]
    }

    competitions = [
        {
            "value": f"{event.name} on {dtStylish(event.start_date, True)} to {dtStylish(event.end_date)}in {event.location}"
        }
        for event in events
        if "deelname op uitnodiging" not in event.notes.lower()
    ]

    competition_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "Select all competitions you want to sign up for:",
                    "questionItem": {
                        "question": {
                            "required": True,
                            "choiceQuestion": {
                                "type": "CHECKBOX",
                                "options": competitions,
                                "shuffle": False
                            }
                        }
                    },
                },
                "location": {
                    "index": 1
                }
            }
        }]
    }

    weightclass_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "What weightclass do you want to compete in?",
                    "questionItem": {
                        "question": {
                            "required": True,
                            "choiceQuestion": {
                                "type": "DROP_DOWN",
                                "options": [
                                    {"value": "F47-"},
                                    {"value": "F52-"},
                                    {"value": "F57-"},
                                    {"value": "F63-"},
                                    {"value": "F69-"},
                                    {"value": "F76-"},
                                    {"value": "F84-"},
                                    {"value": "F84+"},
                                    {"value": "M59-"},
                                    {"value": "M66-"},
                                    {"value": "M74-"},
                                    {"value": "M83-"},
                                    {"value": "M93-"},
                                    {"value": "M105-"},
                                    {"value": "M120-"},
                                    {"value": "M120+"},
                                ],
                                "shuffle": False
                            }
                        }
                    },
                },
                "location": {
                    "index": 2
                }
            }
        }]
    }

    squat_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "Enter an estimate for your potential squat for this competition. This is used to determine the order of the lifters.",
                    "questionItem": {
                        "question": {
                            "required": True,
                            "textQuestion": {}
                        }
                    },
                },
                "location": {
                    "index": 3
                }
            }
        }]
    }

    bench_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "Enter an estimate for your potential bench press for this competition. This is used to determine the order of the lifters.",
                    "questionItem": {
                        "question": {
                            "required": True,
                            "textQuestion": {}
                        }
                    },
                },
                "location": {
                    "index": 4
                }
            }
        }]
    }

    deadlift_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "Enter an estimate for your potential deadlift for this competition. This is used to determine the order of the lifters.",
                    "questionItem": {
                        "question": {
                            "required": True,
                            "textQuestion": {}
                        }
                    },
                },
                "location": {
                    "index": 5
                }
            }
        }]
    }

    coach_name_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "Enter your coaches name if you have one already.",
                    "questionItem": {
                        "question": {
                            "required": False,
                            "textQuestion": {}
                        }
                    },
                },
                "location": {
                    "index": 6
                }
            }
        }]
    }

    coach_knkf_number_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "Enter your coaches KNKF number if you have one already.",
                    "questionItem": {
                        "question": {
                            "required": False,
                            "textQuestion": {}
                        }
                    },
                },
                "location": {
                    "index": 7
                }
            }
        }]
    }

    coach_dob_question = {
        "requests": [{
            "createItem": {
                "item": {
                    "title": "Enter your coaches date of birth if you have one already.",
                    "questionItem": {
                        "question": {
                            "required": False,
                            "dateQuestion": {
                                "includeTime": False,
                                "includeYear": True
                            }
                        }
                    },
                },
                "location": {
                    "index": 8
                }
            }
        }]
    }

    # Add request to remove the initial question
    remove_question = {
        "requests": [{
            "deleteItem": {
                "location": {
                    "index": 9
                }
            }
        }]
    }

    try:
        # Create the Google Form file in the specified folder
        form_file = drive_service.files().create(body=NEW_FORM, media_body=None).execute()
        
        form_id = form_file['id']
        
        new_questions=[update_description, name_question, competition_question, weightclass_question, 
                    squat_question, bench_question, deadlift_question, coach_name_question, coach_knkf_number_question, coach_dob_question, remove_question]
        # Add your code to add questions to the form (similar to your previous code)
        print(f"Form with form ID {form_id} created in folder with ID: {folder_id}")
    
        # Adds the question to the form
        for new_question in new_questions:
            question_setting = forms_service.forms().batchUpdate(formId=form_id, body=new_question).execute()

    except HttpError as error:
        print(f"An HTTP error occurred: {error}")

    # The following request runs a script in google apps scripts API.
    # For now, it changes two things: 
    # Set RequireLogin to False so that anyone can respond to the form and not just users in the organization.
    # Set ShowLinkToRespondAgain to True so that you can easily respond to the form multiple times.

    # Create an execution request object.
    request = {
            'function': 'updateFormSettings',
            'parameters': [form_id]
        }

    try:
        # Make the API request.
        response = scripts_service.scripts().run(scriptId=script_id,
                                            body=request).execute()

    except HttpError as error:
        # The API encountered a problem before the script started executing.
        print(f"An error occurred: {error}")
        print(error.content)
            
    # Load the JSON file into a dictionary
    with open('config.json', 'r') as file:
        config = json.load(file)

    # Edit a single variable in the dictionary
    config['last_access'] = f"{datetime.now().isoformat()}"
    config['form_id'] = form_id
    # Store the order of the questions by id, as apparently the responses are not stored in the same order as the questions, but do contain the questionid. 
    # After retrieving the responses, the answers can be organized by this id order. 
    form = forms_service.forms().get(formId=form_id).execute()
    config['question_order'] = ' '.join([subsubitem["questionId"] for subsubitem in [subitem["question"] for subitem in [item["questionItem"] for item in form["items"]]]])

    # Save the updated dictionary back to the JSON file
    with open('config.json', 'w') as file:
        json.dump(config, file)     
    
def make_form_and_save(new_events):
    create_new_form(new_events)
    save_events_to_file(new_events)

class CompLists(ttk.PanedWindow):
    def __init__(self, parent):
        super().__init__(parent, orient=tk.HORIZONTAL)
        self.root = parent
        self.pane_1 = ttk.Frame(self, padding=(0, 0, 0, 10))
        self.pane_2 = ttk.Frame(self, padding=(0, 0, 0, 10))
        self.pane_3 = ttk.Frame(self, padding=(0, 10, 5, 0))
        self.add(self.pane_1, weight=3)
        self.add(self.pane_2, weight=3)
        self.add(self.pane_3, weight=1)
        self.add_widgets()

    def add_widgets(self):
        # Populate the left and middle listboxes with Event objects
        self.scrollbar = ttk.Scrollbar(self.pane_1)
        self.scrollbar.pack(side="right", fill="y")
        self.tree_left = ttk.Treeview(
            self.pane_1,
            height=11,
            selectmode="browse",
            yscrollcommand=self.scrollbar.set,
        )
        self.scrollbar.config(command=self.tree_left.yview)
        self.tree_left.pack(expand=True, fill="both")
        self.tree_left.column("#0", anchor="w", width=350)
        self.tree_left.heading("#0", text="Known competitions")
        self.tree_left.bind("<<TreeviewSelect>>", self.on_left_tree_select)
        self.old_events = load_previous_events_from_file()
        for item in self.old_events:
            self.tree_left.insert(parent="", index="end", text=item.name)
        
        self.scrollbar = ttk.Scrollbar(self.pane_2)
        self.scrollbar.pack(side="right", fill="y")

        self.tree_right = ttk.Treeview(
            self.pane_2,
            height=11,
            selectmode="browse",
            yscrollcommand=self.scrollbar.set,
        )
        self.scrollbar.config(command=self.tree_right.yview)

        self.tree_right.pack(expand=True, fill="both")
        
        self.tree_right.column("#0", anchor="w", width=350)
        self.tree_right.heading("#0", text="Current Competitions")
        self.new_events = find_events()
        self.tree_right.tag_configure('bg', background='yellow')
        self.tree_right.bind("<<TreeviewSelect>>", self.on_right_tree_select)
        for item in self.new_events:
            self.tree_right.insert(parent="", index="end", text=item.name)

        self.button_whatsapp = ttk.Button(
            self.pane_3, text="Write whatsapp message", style="Switch.info.TButton", command=lambda:write_wa_msg(self.new_events)
        )
        self.button_whatsapp.grid(row=2, column=0, columnspan=2, pady=10)
        self.button_retrieve = ttk.Button(
            self.pane_3, text="Retrieve form responses", style="Switch.info.TButton", command=update_sheets
        )
        self.button_retrieve.grid(row=3, column=0, columnspan=2, pady=10)
        self.button_make_form = ttk.Button(
            self.pane_3, text="Generate new form", style="Switch.info.TButton", command=lambda:make_form_and_save(self.new_events)
        )
        self.button_make_form.grid(row=4, column=0, columnspan=2, pady=10)

    def on_left_tree_select(self, event):
        selected_event = self.old_events[int(self.tree_left.selection()[0][1:])-1]
        tree = self.root.winfo_children()[1].tree
        attribute_order = ["name","signup_date","start_date","end_date","location","organisation","link","notes"]
        for count, item in enumerate(tree.get_children()):
            tree.set(item, column=1, value=getattr(selected_event, attribute_order[count]))
            if tree.item(item)['values'][0] != tree.item(item)['values'][1] != "":
                tree.item(item, tags=("red"))
            else:
                tree.item(item, tags=("white"))
    def on_right_tree_select(self, event):
        selected_event = self.new_events[int(self.tree_right.selection()[0][1:])-1]
        tree = self.root.winfo_children()[1].tree
        attribute_order = ["name","signup_date","start_date","end_date","location","organisation","link","notes"]
        for count, item in enumerate(tree.get_children()):
            tree.set(item, column=2, value=getattr(selected_event, attribute_order[count]))
            if tree.item(item)['values'][1] != tree.item(item)['values'][0] != "":
                tree.item(item, tags=("red"))
            else:
                tree.item(item, tags=("white"))
            
class CompComparator(ttk.PanedWindow):
    def __init__(self, parent):
        super().__init__(parent, orient=tk.HORIZONTAL)

        self.pane_1 = ttk.Frame(self, padding=(0, 0, 0, 10))
        self.pane_2 = ttk.Frame(self, padding=(0, 0, 0, 10))
        self.add(self.pane_1, weight=1)
        self.add(self.pane_2, weight=1)
        self.add_widgets()

    def add_widgets(self):
        self.notebook = ttk.Notebook(self.pane_1)
        self.notebook.pack(expand=True, fill="both")
        self.tab_1 = ttk.Frame(self.notebook)
        self.tab_2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_1, text="Compare Competitions")
        self.notebook.add(self.tab_2, text="Whatsapp Message")

        for index in range(2):
            self.tab_1.columnconfigure(index, weight=1)
            self.tab_1.rowconfigure(index, weight=1)
            
        # Populate the left and middle listboxes with Event objects
        self.tree = ttk.Treeview(
            self.tab_1,
            columns=(1, 2),
            show="tree",
            height=11,
            selectmode="browse",
        )
        self.tree.pack(expand=True, fill="both")
        self.tree.column("#0", anchor="w", width=400)
        self.tree.heading("#0", text="Known competitions")
        self.tree.tag_configure("red", foreground="red")
        self.tree.tag_configure("white", foreground="white")
        self.old_events = load_previous_events_from_file()
        # Create labels and display attributes
        attributes = ["Name","Signup Date","Start Date",
                      "End Date","Location","Organisation",
                      "Link", "Notes"
        ]
        for attribute in attributes:
            self.tree.insert(parent="", index="end", text=attribute)
        self.tree.column("#0", anchor="w", width=150)
        self.tree.column(1, anchor="w", width=400)
        self.tree.column(2, anchor="w", width=400)
        
        self.reader = Reader(self.tab_2)
        self.reader.pack(fill='both', expand='yes')

class Reader(ttk.Frame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.configure(padding=10)
        self.filename = tk.StringVar()

        # scrolled text with custom highlight colors
        self.text_area = ScrolledText(self)
        self.text_area.pack(fill='both')

        path = "whatsapp_msg.txt"
        with open(path, encoding='utf-8') as f:
            self.text_area.delete('1.0', 'end')
            self.text_area.insert('end', f.read())
            self.filename.set(path)



class App(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=15)
        pane1 = CompLists(self)
        pane1.grid(row=0, column=0, rowspan=1)
        pane2 = CompComparator(self)
        pane2.grid(row=1, column=0, rowspan=1)
    


def main():
    windll.shcore.SetProcessDpiAwareness(1)
    root = tk.Tk()
    root.title("Competition Signup")
    sv_ttk.set_theme("dark")
    App(root).pack(expand=True, fill="both")
    root.mainloop()


if __name__ == "__main__":
    main()