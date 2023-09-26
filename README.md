# The project
This project was started to automate some work that is needed for the student association for powerlifting and weightlifting. The secretary is tasked with collecting signups from the members who wish to compete. 
Until now, this has been managed manually. This manual work consists of:
1. Regularly checking the KNKF website for new competitions
2. Creating a google form with the new competitions and sharing it with members
3. Checking if the members who have signed up are allowed to join the particular competition
4. Checking if members have paid the competition fee
5. Entering the info given by members when the competition is open for signup.

So far, tasks 2 and 3 are performed by the script: application.py

# Installation
At the moment, the script still needs to be run inside an IDE like Visual Studio Code. 
1. Create a virtual environment and run ```pip install -r requirements.txt```.
2. Create a google oauth client ID, see: [info](https://support.google.com/cloud/answer/6158849?hl=en)
3. Enable the Google Sheets API, the Gmail API, and the Google Drive API, see: [info](https://support.google.com/googleapi/answer/6158841?hl=en)
4. Create a new script in Google Apps Script and copy paste the code in [form_opener](form_opener.js) and deploy it as an API executable. Note the script id for the next step.
5. Configure the required variables [folder_id, script_id and sheet_id] in data/saved_vars.json, see below for an overview.
6. Start application.py
7. The application will now request access to the scopes required by the Google API.
Note: You need to login with a gmail account that receives association invites from the KNKF. Other mail clients are not supported (yet).

# Use
When the script is run, it automatically scrapes the KNKF page for new competitions. The top left window shows the competitions that were included in the previous form. The window to the right of it shows the competitions currently listed on the KNKF page. You can compare the old competitions with the new competitions to see if any changes have been made and decide if a new form should be generated. 
* Clicking 'retreive responses' will retrieve all new responses of the form into the sheet, and append info about their eligibility to compete. 
* Clicking 'generate new form' will generate a new form with the info contained in the 'Current Competitions' list. To be sure no data is lost, the old form is not deleted automatically. 
* Clicking 'write whatsapp message' will write the 'Current Competitions' into a nice message with the currently known form included. The message is not sent automatically to allow a manual check first. 

# Saved vars
* saved_vars.json contains the following variables:
* form_id: the id of the currently active Google form. This id is also present in the URL when opening the form. 
* last_access: the last time the form has been accessed by this script. This is used to filter out old responses.
* question_order: the questions in the form contain an id. This variable contains the order of the question ids by which the responses containing that same id can be sorted. For some reason, the responses do not contain the * same order as the questions in the form do, so this is my solution. 
* fuzzy_threshold: The threshold used for comparing names. Members enter their name in the form, and the script checks to see if that name is known or not. The fuzzywuzzy package is used to find the most similar name in the * database, and the threshold is then used to reject names which are not similar enough.
* month_mapping: This is just here because it is a messy list to map month abbreviations with their full version. 
* folder_id: This is the id associated with the Google Drive folder where automatically generated forms are stored.
* script_id: This is the id associated with the Google Apps script which opens a form for members which are not registered with the Google organisation IJzersterk. 
* sheet_id: This is the id associated with the sheet where new responses are stored.



![Image of the GUI](GUI_screenshot.png?raw=true "Title")


# WARNING
Be careful with the files in the tokens folder. These contain authentication keys which allow access to your google account.
