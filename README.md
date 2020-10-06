# my_schedule

Small script, that reads the schedule of today or tomorrow from my Office365 account and returns it as a markdown list, 
which I can import into Roam Research.

## Installation

Create a virtual environment and install the dependencies
```bash
python3 -m venv .venv        
source .venv/bin/activate
pip install -r requirements.txt
```

## Configuration

### OAuth Client Application
Follow the steps decribed here: [O365/python-o365: A simple python library to interact with Microsoft Graph and Office 365 API](https://github.com/O365/python-o365#authentication-steps)

### App Config
Copy the `credentials_template.py` file to `credentials.py` and replace the placeholders with the `client id` and the 
`client secret` from the step before. 

## Running the Script

Make sure your virtual environment.

First you have to logon by running: `python my_schedule logon`. Just follow the instructions on the screen.

Then you can get the schedule by using `python my_schedule today` or `python my_schedule tomorrow`

`python my_schedule --help` shows a short help.
