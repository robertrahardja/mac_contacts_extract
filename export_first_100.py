#!/usr/bin/env python3
"""
Export first 100 contacts to test the setup
"""

import os
import sys
import subprocess
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Google Sheets imports
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

# Google Sheets API scope
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def export_first_contacts():
    """Export just the first 100 contacts as a test"""
    print("📱 Exporting first 100 contacts as a test...")

    applescript = '''
    tell application "Contacts"
        set output to ""
        set allPeople to people 1 through 100

        repeat with aPerson in allPeople
            -- Get name
            try
                set firstName to first name of aPerson
            on error
                set firstName to ""
            end try

            try
                set lastName to last name of aPerson
            on error
                set lastName to ""
            end try

            -- Get primary email
            try
                set emailList to emails of aPerson
                if (count of emailList) > 0 then
                    set primaryEmail to value of item 1 of emailList
                else
                    set primaryEmail to ""
                end if
            on error
                set primaryEmail to ""
            end try

            -- Get primary phone
            try
                set phoneList to phones of aPerson
                if (count of phoneList) > 0 then
                    set primaryPhone to value of item 1 of phoneList
                else
                    set primaryPhone to ""
                end if
            on error
                set primaryPhone to ""
            end try

            -- Create tab-separated line
            set personData to firstName & tab & lastName & tab & primaryEmail & tab & primaryPhone
            set output to output & personData & linefeed
        end repeat

        return output
    end tell
    '''

    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            timeout=30
        )

        if result.returncode != 0:
            print(f"❌ Error: {result.stderr}")
            return None

        # Parse the output
        contacts_data = []
        lines = result.stdout.strip().split('\n')

        for line in lines:
            if line:
                parts = line.split('\t')
                if len(parts) >= 4:
                    contact = {
                        'First Name': parts[0],
                        'Last Name': parts[1],
                        'Email': parts[2],
                        'Phone': parts[3]
                    }
                    # Skip empty contacts
                    if any(parts):
                        contacts_data.append(contact)

        print(f"✅ Exported {len(contacts_data)} contacts")
        return contacts_data

    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def authenticate_google_sheets():
    """Authenticate and return Google Sheets service"""
    creds = None
    token_file = 'token.json'

    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("❌ credentials.json not found!")
                sys.exit(1)

            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)

    return build('sheets', 'v4', credentials=creds)

def update_sheet(service, contacts_data):
    """Update Google Sheet"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found!")
        return None

    print(f"📊 Updating Google Sheet...")

    if not contacts_data:
        return None

    # Prepare data
    headers = ['First Name', 'Last Name', 'Email', 'Phone']
    values = [headers]
    for contact in contacts_data:
        row = [contact.get(h, '') for h in headers]
        values.append(row)

    try:
        # Clear and update
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range="Contacts!A:D"
        ).execute()

        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range="Contacts!A1",
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        print(f"✅ Updated {len(contacts_data)} contacts")
        print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return sheet_id

    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def main():
    print("=" * 50)
    print("Test Export: First 100 Contacts")
    print("=" * 50)

    # Export contacts
    contacts = export_first_contacts()
    if not contacts:
        print("Failed to export contacts")
        sys.exit(1)

    # Authenticate
    print("\n🔐 Authenticating...")
    service = authenticate_google_sheets()

    # Update sheet
    if update_sheet(service, contacts):
        print("\n✨ Test successful!")
        print("Check your Google Sheet to see the first 100 contacts.")
        print("\nNote: Due to the large number of contacts (3,792),")
        print("you may need to export them in smaller groups or")
        print("use the AppleScript version for a CSV export.")

if __name__ == "__main__":
    main()