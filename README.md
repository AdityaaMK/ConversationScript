## ConversationScript
Reads a qualtrics form and fills out conversation data with residents.
CURRENTLY ONLY WORKS FOR NORTH AVE AREA AND BROWN AND HARRIS

# Update
I have found a better way to install the chrome driver using with using the python package manager. Please see below to find new installation instructions.

# Setup and Usage

**How to download**  
On this page, find the code button and then either download the zip file or use git clone in your terminal to place the repository where you please.

**Fill out Conversations Spreadsheet**
Make sure to follow the instructions on the spreadsheet and then delete row 2(instructions).

**Initial setup- Only do this one time**  
1. Navigate to the ConversationScript directory
2. Create a python virtual environment: Run `python3 -m venv ./ra-venv` This will create the virtual environment called ra-venv.
3. Activate your virtual environment: Run `. ./ra-venv/bin/activate`
4. Install dependencies: `pip install -r requirements.txt`

**Usage - do this every time**  
1. Navigate to the ConversationScript directory
2. Activate your virtual environment: `. ./ra-venv/bin/activate`
3. Fill out spreadsheet
4. Run `python3 convo_script.py`

# Modifications
The only two files that you would need to modify is the convo_script.py or
the excel_reader.py and the locations are all marked with the comment of "#TODO"

**In convoscript.py**  
- want_email in the convo_script.py is personal preference. Set this value to true if you would like to recieve a confirmation email, but make sure you actually have an email in the corresponding column of your spreadsheet otherwise it will send the email to Kevin :)...

**In excel_reader.py**  
- loc in excel_reader will change where the directory is for the excel file containing the list of residents and information


