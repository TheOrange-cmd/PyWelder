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
        
class ThemedGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Themed GUI")
        self.root.geometry("600x400")

        # Create left column listbox
        self.left_frame = ttk.Frame(self.root)
        self.left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ns")
        self.left_listbox = tk.Listbox(self.left_frame, selectmode=tk.SINGLE, exportselection=0)
        self.left_listbox.pack(fill=tk.BOTH, expand=True)
        self.left_listbox.bind("<<ListboxSelect>>", self.on_select)

        # Create middle column listbox
        self.middle_frame = ttk.Frame(self.root)
        self.middle_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ns")
        self.middle_listbox = tk.Listbox(self.middle_frame, selectmode=tk.SINGLE, exportselection=0)
        self.middle_listbox.pack(fill=tk.BOTH, expand=True)
        self.middle_listbox.bind("<<ListboxSelect>>", self.on_select)

        # Create right column with buttons
        self.right_frame = ttk.Frame(self.root)
        self.right_frame.grid(row=0, column=2, padx=10, pady=10, sticky="ns")
        self.button1 = ttk.Button(self.right_frame, text="Compare", command=self.compare_events)
        self.button1.pack(fill=tk.BOTH, pady=5)
        self.button2 = ttk.Button(self.right_frame, text="Button 2", command=self.on_button2_click)
        self.button2.pack(fill=tk.BOTH, pady=5)
        self.button3 = ttk.Button(self.right_frame, text="Button 3", command=self.on_button3_click)
        self.button3.pack(fill=tk.BOTH, pady=5)

        # Initialize selected events as None
        self.selected_event_left = None
        self.selected_event_right = None

        # Populate the left and middle listboxes with sample Event objects
        self.old_events = load_previous_events_from_file()
        self.new_events = find_events()
        for event in self.old_events:
            self.left_listbox.insert(tk.END, event.name)
        for event in self.new_events:
            self.middle_listbox.insert(tk.END, event.name)

    def on_select(self, event):
        sender = event.widget
        selected_index = sender.curselection()
        if selected_index:
            selected_index = int(selected_index[0])
            selected_event_name = sender.get(selected_index)
            

            if sender == self.left_listbox:
                selected_event = next((event for event in self.old_events if event.name == selected_event_name), None)
                self.selected_event_left = selected_event
            elif sender == self.middle_listbox:
                selected_event = next((event for event in self.new_events if event.name == selected_event_name), None)
                self.selected_event_right = selected_event

    def compare_events(self):
        if self.selected_event_left and self.selected_event_right:
            self.display_event_attributes(self.selected_event_left, self.selected_event_right)
        else:
            messagebox.showwarning("Selection Error", "Please select events from both lists to compare.")

    def compare_attribute(self, attr1, attr2):
        # Compare two attribute values using Fuzzy Wuzzy and a threshold of 80
        similarity_ratio = fuzz.ratio(attr1, attr2)
        return similarity_ratio >= 95
    
    def display_event_attributes(self, event1, event2):
        attributes_window = tk.Toplevel(self.root)
        attributes_window.title("Event Attributes")

        attributes_frame = ttk.Frame(attributes_window)
        attributes_frame.pack(padx=10, pady=10)

        # Create labels and display attributes
        attributes = [
            ("Name", event1.name, event2.name),
            ("Signup Date", event1.signup_date.strftime("%Y-%m-%d"), event2.signup_date.strftime("%Y-%m-%d")),
            ("Start Date", event1.start_date.strftime("%Y-%m-%d"), event2.start_date.strftime("%Y-%m-%d")),
            ("End Date", event1.end_date.strftime("%Y-%m-%d"), event2.end_date.strftime("%Y-%m-%d")),
            ("Location", event1.location, event2.location),
            ("Organisation", event1.organisation, event2.organisation),
            ("Link", event1.link, event2.link),
            ("Notes", event1.notes, event2.notes)
        ]

        for i, (label, value1, value2) in enumerate(attributes):
            ttk.Label(attributes_frame, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=5)
          # Compare attribute values and highlight if not similar enough 
            if self.compare_attribute(value1, value2):
                fg = "green"
            else:
                fg = "red"
            ttk.Label(attributes_frame, text=value1, background=fg).grid(row=i, column=1, padx=5, pady=5)
            ttk.Label(attributes_frame, text=value2, background=fg).grid(row=i, column=2, padx=5, pady=5)


    def on_button2_click(self):
        print("Button 2 clicked")

    def on_button3_click(self):
        print("Button 3 clicked")
if __name__ == "__main__":
    root = tk.Tk()
    sv_ttk.set_theme("dark")
    app = ThemedGUI(root)
    
    root.mainloop()
