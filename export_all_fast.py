#!/usr/bin/env python3
"""
FAST Mac Contacts Export - Optimized for speed
Exports all contacts with minimal AppleScript calls
"""

import os
import sys
import subprocess
import json
import time
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

def get_all_contacts_batch():
    """Get ALL contacts data in one massive AppleScript call"""
    print("📦 Fetching all contacts in a single batch...")
    print("⏳ This will take 1-2 minutes for 3,792 contacts...")

    script = '''
    tell application "Contacts"
        set output to ""
        set contactCount to count of people

        repeat with i from 1 to contactCount
            set p to person i
            set contactData to ""

            -- Basic fields
            try
                set contactData to contactData & (first name of p)
            end try
            set contactData to contactData & "|"

            try
                set contactData to contactData & (last name of p)
            end try
            set contactData to contactData & "|"

            try
                set contactData to contactData & (organization of p)
            end try
            set contactData to contactData & "|"

            -- Emails (simplified - just get first of each type)
            set homeEmail to ""
            set workEmail to ""
            try
                set emailList to emails of p
                repeat with em in emailList
                    set emLabel to ""
                    try
                        set emLabel to label of em as string
                    end try
                    if emLabel contains "Work" and workEmail is "" then
                        set workEmail to value of em
                    else if emLabel contains "Home" and homeEmail is "" then
                        set homeEmail to value of em
                    end if
                end repeat
            end try
            set contactData to contactData & homeEmail & "|" & workEmail & "|"

            -- Phones (simplified - just get first of each type)
            set mobilePhone to ""
            set workPhone to ""
            set homePhone to ""
            try
                set phoneList to phones of p
                repeat with ph in phoneList
                    set phLabel to ""
                    try
                        set phLabel to label of ph as string
                    end try
                    if (phLabel contains "Mobile" or phLabel contains "iPhone") and mobilePhone is "" then
                        set mobilePhone to value of ph
                    else if phLabel contains "Work" and not phLabel contains "FAX" and workPhone is "" then
                        set workPhone to value of ph
                    else if phLabel contains "Home" and not phLabel contains "FAX" and homePhone is "" then
                        set homePhone to value of ph
                    end if
                end repeat
            end try
            set contactData to contactData & mobilePhone & "|" & workPhone & "|" & homePhone & "|"

            -- Notes
            try
                set noteText to note of p
                if noteText is not missing value then
                    -- Replace newlines to avoid breaking our format
                    set AppleScript's text item delimiters to return
                    set noteItems to text items of noteText
                    set AppleScript's text item delimiters to " "
                    set noteText to noteItems as string
                    set contactData to contactData & noteText
                end if
            end try

            -- Add contact if it has any data
            if contactData is not "||||||||" then
                set output to output & contactData & "\\n"
            end if

            -- Progress indicator every 100 contacts
            if i mod 100 = 0 then
                log "Processed " & i & " contacts..."
            end if
        end repeat

        return output
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script],
                              capture_output=True, text=True, timeout=300)
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            contacts = []

            for line in lines:
                if line:
                    parts = line.split('|')
                    if len(parts) >= 9:
                        contact = {
                            'First Name': parts[0],
                            'Last Name': parts[1],
                            'Organization': parts[2],
                            'Home Email': parts[3],
                            'Work Email': parts[4],
                            'Mobile Phone': parts[5],
                            'Work Phone': parts[6],
                            'Home Phone': parts[7],
                            'Notes': parts[8] if len(parts) > 8 else ''
                        }
                        # Only add if contact has some data
                        if any(contact.values()):
                            contacts.append(contact)

            print(f"✅ Extracted {len(contacts)} contacts")
            return contacts
        else:
            print(f"❌ AppleScript error: {result.stderr}")
            return []
    except subprocess.TimeoutExpired:
        print("❌ Export timed out after 5 minutes")
        return []
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

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

def upload_to_sheets(service, contacts):
    """Upload contacts to Google Sheets"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')

    if not contacts:
        print("❌ No contacts to upload")
        return False

    print(f"\n📊 Uploading {len(contacts)} contacts to Google Sheets...")

    # Headers
    headers = [
        'First Name', 'Last Name', 'Organization',
        'Home Email', 'Work Email',
        'Mobile Phone', 'Work Phone', 'Home Phone',
        'Notes'
    ]

    # Build values array
    values = [headers]
    for contact in contacts:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear sheet first
        print("   Clearing existing data...")
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Upload all data at once
        print(f"   Uploading {len(values)-1} contacts...")
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        # Format headers
        print("   Formatting sheet...")
        requests = [
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': 0,
                        'gridProperties': {'frozenRowCount': 1}
                    },
                    'fields': 'gridProperties.frozenRowCount'
                }
            },
            {
                'repeatCell': {
                    'range': {'sheetId': 0, 'startRowIndex': 0, 'endRowIndex': 1},
                    'cell': {'userEnteredFormat': {'textFormat': {'bold': True}}},
                    'fields': 'userEnteredFormat.textFormat.bold'
                }
            },
            {
                'autoResizeDimensions': {
                    'dimensions': {
                        'sheetId': 0,
                        'dimension': 'COLUMNS',
                        'startIndex': 0,
                        'endIndex': len(headers)
                    }
                }
            }
        ]

        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={'requests': requests}
        ).execute()

        print(f"✅ Successfully uploaded {len(contacts)} contacts!")
        print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return True

    except Exception as e:
        print(f"❌ Upload error: {e}")
        return False

def save_backup(contacts):
    """Save backup JSON file"""
    if not contacts:
        return None

    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"fast_export_{timestamp}.json"

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts, f, indent=2, ensure_ascii=False)

    print(f"💾 Backup saved: {json_file}")
    return json_file

def main():
    """Main function"""
    print("="*60)
    print("FAST MAC CONTACTS EXPORT")
    print("Optimized for 3,792 contacts")
    print("="*60)

    start_time = time.time()

    # Get all contacts in one batch
    contacts = get_all_contacts_batch()

    if not contacts:
        print("❌ No contacts exported")
        sys.exit(1)

    # Save backup
    save_backup(contacts)

    # Upload to Google Sheets
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    if upload_to_sheets(service, contacts):
        elapsed = time.time() - start_time
        print(f"\n🎉 EXPORT COMPLETE!")
        print(f"⏱️ Total time: {elapsed:.1f} seconds")
        print(f"📊 {len(contacts)} contacts exported")
    else:
        print("\n❌ Upload failed")

if __name__ == "__main__":
    main()