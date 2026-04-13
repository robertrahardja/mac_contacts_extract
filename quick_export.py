#!/usr/bin/env python3
"""
Quick Mac Contacts Export to Google Sheets
Simplified version that exports only essential fields
"""

import os
import sys
import subprocess
import json
from datetime import datetime
from pathlib import Path
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

def export_contacts_simple():
    """Export contacts with only essential fields"""
    print("📱 Exporting contacts from Mac Contacts app (simplified)...")

    # Simple AppleScript to get just names and primary email/phone
    applescript = '''
    tell application "Contacts"
        set output to ""
        set allPeople to every person
        set totalCount to count of allPeople

        -- Limit to first 1000 contacts for speed
        if totalCount > 1000 then
            set peopleToProcess to items 1 through 1000 of allPeople
        else
            set peopleToProcess to allPeople
        end if

        repeat with aPerson in peopleToProcess
            set personData to ""

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

            -- Get organization
            try
                set org to organization of aPerson
            on error
                set org to ""
            end try

            -- Create tab-separated line
            set personData to firstName & tab & lastName & tab & primaryEmail & tab & primaryPhone & tab & org
            set output to output & personData & linefeed
        end repeat

        return output
    end tell
    '''

    try:
        # Run AppleScript with longer timeout
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            timeout=60  # 60 second timeout
        )

        if result.returncode != 0:
            print(f"❌ AppleScript error: {result.stderr}")
            return None

        # Parse the tab-separated output
        contacts_data = []
        lines = result.stdout.strip().split('\n')

        for line in lines:
            if line:
                parts = line.split('\t')
                if len(parts) >= 5:
                    contact = {
                        'First Name': parts[0],
                        'Last Name': parts[1],
                        'Email': parts[2],
                        'Phone': parts[3],
                        'Organization': parts[4] if len(parts) > 4 else ''
                    }
                    contacts_data.append(contact)

        print(f"✅ Successfully exported {len(contacts_data)} contacts")
        return contacts_data

    except subprocess.TimeoutExpired:
        print("❌ Export timed out. Your contact list may be too large.")
        print("   Try reducing the number of contacts in your Contacts app.")
        return None
    except Exception as e:
        print(f"❌ Error during export: {e}")
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

    service = build('sheets', 'v4', credentials=creds)
    return service

def update_google_sheet(service, contacts_data):
    """Update Google Sheet with contacts data"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found in .env file")
        sys.exit(1)

    sheet_name = os.getenv('SHEET_NAME', 'Contacts')
    print(f"📊 Updating Google Sheet...")

    if not contacts_data:
        print("No contacts to export")
        return None

    # Prepare headers and values
    headers = list(contacts_data[0].keys())
    values = [headers]
    for contact in contacts_data:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear existing content
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Update with new data
        body = {'values': values}
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body=body
        ).execute()

        # Simple formatting
        requests = [
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': 0,
                        'gridProperties': {
                            'frozenRowCount': 1
                        }
                    },
                    'fields': 'gridProperties.frozenRowCount'
                }
            },
            {
                'repeatCell': {
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': 0,
                        'endRowIndex': 1
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.textFormat.bold'
                }
            }
        ]

        batch_update_body = {'requests': requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body=batch_update_body
        ).execute()

        print(f"✅ Updated {result.get('updatedCells')} cells")
        print(f"📋 Sheet URL: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return sheet_id

    except Exception as e:
        print(f"❌ Error updating sheet: {str(e)}")
        return None

def save_backup(contacts_data):
    """Save local backup"""
    if not contacts_data:
        return

    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"contacts_simple_{timestamp}.json"

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts_data, f, indent=2, ensure_ascii=False)

    print(f"💾 Backup saved: {json_file}")
    return json_file

def main():
    """Main function"""
    print("=" * 50)
    print("Quick Mac Contacts Export to Google Sheets")
    print("=" * 50)

    # Export contacts
    contacts = export_contacts_simple()

    if not contacts:
        print("❌ Export failed")
        sys.exit(1)

    # Save backup
    save_backup(contacts)

    # Authenticate with Google
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    # Update Google Sheet
    sheet_id = update_google_sheet(service, contacts)

    if sheet_id:
        print("\n✨ Export complete!")
        print(f"📱 Total contacts exported: {len(contacts)}")
        print(f"📊 View your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
    else:
        print("\n⚠️ Export to Google Sheets failed, but backup was saved")

if __name__ == "__main__":
    main()