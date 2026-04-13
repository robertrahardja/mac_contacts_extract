#!/usr/bin/env python3
"""
Simple test with 3 contacts - basic fields only
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

def export_contact_simple(index):
    """Export single contact with basic fields only"""

    applescript = f'''
    tell application "Contacts"
        try
            set p to person {index}

            -- Names
            try
                set fn to first name of p
                if fn is missing value then set fn to ""
            on error
                set fn to ""
            end try

            try
                set ln to last name of p
                if ln is missing value then set ln to ""
            on error
                set ln to ""
            end try

            try
                set org to organization of p
                if org is missing value then set org to ""
            on error
                set org to ""
            end try

            -- Primary email
            set emailStr to ""
            try
                set emailList to emails of p
                if (count of emailList) > 0 then
                    set emailStr to value of item 1 of emailList
                end if
            on error
                set emailStr to ""
            end try

            -- Primary phone
            set phoneStr to ""
            try
                set phoneList to phones of p
                if (count of phoneList) > 0 then
                    set phoneStr to value of item 1 of phoneList
                end if
            on error
                set phoneStr to ""
            end try

            -- Notes
            try
                set nt to note of p
                if nt is missing value then set nt to ""
            on error
                set nt to ""
            end try

            return fn & "|" & ln & "|" & org & "|" & emailStr & "|" & phoneStr & "|" & nt

        on error errMsg
            return "ERROR: " & errMsg
        end try
    end tell
    '''

    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            timeout=5
        )

        if result.returncode == 0 and not result.stdout.strip().startswith("ERROR"):
            return result.stdout.strip()
        else:
            print(f"Error for contact {index}: {result.stderr}")
            return None
    except Exception as e:
        print(f"Exception for contact {index}: {e}")
        return None

def export_3_contacts_simple():
    """Export first 3 contacts with basic fields"""
    print("🧪 Testing simple export with first 3 contacts...")

    contacts_data = []

    for i in range(1, 4):
        print(f"   Exporting contact {i}...")

        contact_data = export_contact_simple(i)

        if contact_data:
            parts = contact_data.split('|')
            print(f"   ✅ Contact {i}: {parts[0]} {parts[1]} - {parts[2]}")

            contact = {
                'First Name': parts[0],
                'Last Name': parts[1],
                'Organization': parts[2],
                'Primary Email': parts[3],
                'Primary Phone': parts[4],
                'Notes': parts[5] if len(parts) > 5 else ''
            }

            contacts_data.append(contact)
        else:
            print(f"   ❌ Contact {i}: Failed")

    print(f"\n✅ Exported {len(contacts_data)} contacts successfully")
    return contacts_data

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
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)

    return build('sheets', 'v4', credentials=creds)

def upload_to_sheets(service, contacts_data):
    """Upload test data to Google Sheets"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Contacts')

    print(f"\n📊 Uploading to Google Sheets...")

    headers = ['First Name', 'Last Name', 'Organization', 'Primary Email', 'Primary Phone', 'Notes']
    values = [headers]

    for contact in contacts_data:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear and upload
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        # Format headers
        requests = [{
            'repeatCell': {
                'range': {'sheetId': 0, 'startRowIndex': 0, 'endRowIndex': 1},
                'cell': {'userEnteredFormat': {'textFormat': {'bold': True}}},
                'fields': 'userEnteredFormat.textFormat.bold'
            }
        }]

        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={'requests': requests}
        ).execute()

        print(f"✅ Success! Uploaded {len(contacts_data)} contacts")
        print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return sheet_id

    except Exception as e:
        print(f"❌ Upload error: {e}")
        return None

def main():
    print("=" * 50)
    print("SIMPLE TEST: 3 Contacts to Google Sheets")
    print("=" * 50)

    # Export 3 contacts
    contacts = export_3_contacts_simple()

    if not contacts:
        print("❌ No contacts exported")
        sys.exit(1)

    # Upload to Google Sheets
    print("\n🔐 Authenticating...")
    service = authenticate_google_sheets()

    sheet_id = upload_to_sheets(service, contacts)

    if sheet_id:
        print("\n🎉 TEST SUCCESSFUL!")
        print("✅ Basic export works perfectly")
        print("✅ Google Sheets upload successful")
        print("\nNext step: Run comprehensive export with all fields")
    else:
        print("\n❌ Test failed")

if __name__ == "__main__":
    main()