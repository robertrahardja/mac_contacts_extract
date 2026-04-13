#!/usr/bin/env python3
"""
Mac Contacts to Google Sheets - Hybrid Exporter
Uses AppleScript to export contacts and Python for Google Sheets upload
"""

import os
import sys
import csv
import json
import subprocess
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

def export_contacts_via_applescript():
    """Export contacts using AppleScript"""
    print("📱 Counting contacts in Mac Contacts app...")

    # First, get the count
    count_script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', count_script], capture_output=True, text=True)
    total_contacts = int(result.stdout.strip())
    print(f"📊 Found {total_contacts} total contacts")

    if total_contacts > 100:
        print(f"⚠️  Large contact list detected ({total_contacts} contacts)")
        print("   Processing in batches to avoid timeout...")

    # Process in smaller batches for large contact lists
    batch_size = 50
    all_contacts = []

    for start_idx in range(0, min(total_contacts, 500), batch_size):  # Limit to first 500 for testing
        end_idx = min(start_idx + batch_size, total_contacts)
        print(f"   Processing contacts {start_idx + 1} to {end_idx}...")

        applescript = f'''
    tell application "Contacts"
        set contactList to {{{{}}}}
        set allPeople to people {start_idx + 1} through {end_idx}

        repeat with aPerson in allPeople
            set contactInfo to {{{{}}}}

            -- Get basic names
            try
                set contactInfo to contactInfo & {{first name of aPerson}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            try
                set contactInfo to contactInfo & {{last name of aPerson}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            try
                set contactInfo to contactInfo & {{middle name of aPerson}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            try
                set contactInfo to contactInfo & {{nickname of aPerson}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            -- Organization
            try
                set contactInfo to contactInfo & {{organization of aPerson}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            try
                set contactInfo to contactInfo & {{job title of aPerson}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            -- Get first email
            try
                set emailList to emails of aPerson
                if (count of emailList) > 0 then
                    set firstEmail to value of item 1 of emailList
                    set contactInfo to contactInfo & {{firstEmail}}
                else
                    set contactInfo to contactInfo & {{""}}
                end if
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            -- Get first phone
            try
                set phoneList to phones of aPerson
                if (count of phoneList) > 0 then
                    set firstPhone to value of item 1 of phoneList
                    set contactInfo to contactInfo & {{firstPhone}}
                else
                    set contactInfo to contactInfo & {{""}}
                end if
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            -- Get birthday
            try
                set bday to birth date of aPerson
                set contactInfo to contactInfo & {{(month of bday as string) & "/" & (day of bday as string) & "/" & (year of bday as string)}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            -- Get note
            try
                set contactInfo to contactInfo & {{note of aPerson}}
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            -- Get first address
            try
                set addressList to addresses of aPerson
                if (count of addressList) > 0 then
                    set addr to item 1 of addressList
                    set fullAddress to ""
                    try
                        set fullAddress to street of addr
                    end try
                    try
                        set fullAddress to fullAddress & ", " & city of addr
                    end try
                    try
                        set fullAddress to fullAddress & ", " & state of addr
                    end try
                    try
                        set fullAddress to fullAddress & " " & zip of addr
                    end try
                    set contactInfo to contactInfo & {{fullAddress}}
                else
                    set contactInfo to contactInfo & {{""}}
                end if
            on error
                set contactInfo to contactInfo & {{""}}
            end try

            set end of contactList to contactInfo
        end repeat

            return contactList
        end tell
        '''

        try:
            # Run AppleScript for this batch
            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=10  # Shorter timeout for smaller batches
            )

            if result.returncode != 0:
                print(f"   ⚠️ Error in batch {start_idx}-{end_idx}: {result.stderr}")
                continue

            # Parse the AppleScript output
            output = result.stdout.strip()

            # Convert AppleScript list format to Python
            if output and output != '{}':
                # Parse the AppleScript list output
                # Remove outer braces and split by contact
                output = output.strip('{}')

                # Simple parsing for the structured output
                current_contact = []
                field_count = 0
                in_quotes = False
                current_field = ""

                for char in output:
                    if char == '"' and (not current_field or current_field[-1] != '\\'):
                        in_quotes = not in_quotes
                    elif char == ',' and not in_quotes:
                        current_contact.append(current_field.strip(' "'))
                        current_field = ""
                        field_count += 1

                        if field_count == 11:  # We have 11 fields per contact
                            all_contacts.append(current_contact)
                            current_contact = []
                            field_count = 0
                    else:
                        current_field += char

                # Don't forget the last field
                if current_field:
                    current_contact.append(current_field.strip(' "'))
                    field_count += 1
                    if field_count == 11:
                        all_contacts.append(current_contact)

        except subprocess.TimeoutExpired:
            print(f"   ⚠️ Batch {start_idx}-{end_idx} timed out, skipping...")
            continue
        except Exception as e:
            print(f"   ⚠️ Error in batch: {e}")
            continue

    print(f"✅ Successfully exported {len(all_contacts)} contacts")
    return all_contacts

def contacts_to_dict_list(contacts):
    """Convert contacts list to dictionary format"""
    if not contacts:
        return []

    headers = [
        'First Name', 'Last Name', 'Middle Name', 'Nickname',
        'Organization', 'Job Title', 'Email', 'Phone',
        'Birthday', 'Note', 'Address'
    ]

    dict_list = []
    for contact in contacts:
        # Ensure we have the right number of fields
        while len(contact) < len(headers):
            contact.append('')

        contact_dict = {headers[i]: contact[i] for i in range(len(headers))}
        dict_list.append(contact_dict)

    return dict_list

def authenticate_google_sheets():
    """Authenticate and return Google Sheets service"""
    creds = None
    token_file = 'token.json'

    # Token file stores the user's access and refresh tokens
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("❌ credentials.json not found!")
                print("Please follow the setup guide to get your Google API credentials")
                sys.exit(1)

            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for next run
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service

def update_google_sheet(service, contacts_data):
    """Update Google Sheet with contacts data"""

    # Get sheet ID from environment
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found in .env file")
        sys.exit(1)

    sheet_name = os.getenv('SHEET_NAME', 'Contacts')

    print(f"📊 Updating Google Sheet: {sheet_id}")

    # Prepare headers
    if contacts_data:
        headers = list(contacts_data[0].keys())
    else:
        print("No contacts found to export")
        return None

    # Prepare values (headers + data rows)
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
        body = {
            'values': values
        }

        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body=body
        ).execute()

        # Format the sheet
        requests = [
            # Freeze header row
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
            # Bold header row
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
            },
            # Auto-resize columns
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
        print("\nMake sure:")
        print("1. The Google Sheet ID in .env is correct")
        print("2. The sheet exists and you have edit access")
        print("3. The Google Sheets API is enabled in your project")
        return None

def save_local_backup(contacts_data):
    """Save local backup of contacts"""
    export_dir = Path(os.getenv('EXPORT_DIR', 'exports'))
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Save as JSON
    json_file = export_dir / f"contacts_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts_data, f, indent=2, ensure_ascii=False)

    # Save as CSV
    csv_file = export_dir / f"contacts_{timestamp}.csv"
    if contacts_data:
        headers = list(contacts_data[0].keys())
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            writer.writerows(contacts_data)

    print(f"💾 Backups saved:")
    print(f"   JSON: {json_file}")
    print(f"   CSV:  {csv_file}")

    return json_file, csv_file

def main():
    """Main function"""
    print("=" * 50)
    print("Mac Contacts to Google Sheets Exporter")
    print("=" * 50)

    # Export contacts using AppleScript
    contacts = export_contacts_via_applescript()

    if not contacts:
        print("❌ Failed to export contacts")
        sys.exit(1)

    # Convert to dictionary format
    contacts_data = contacts_to_dict_list(contacts)

    if not contacts_data:
        print("No contacts found in Contacts app")
        return

    # Save local backup
    json_file, csv_file = save_local_backup(contacts_data)

    # Authenticate with Google
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    # Update Google Sheet
    sheet_id = update_google_sheet(service, contacts_data)

    if sheet_id:
        print("\n✨ Export complete!")
        print(f"📱 Total contacts exported: {len(contacts_data)}")
        print(f"📊 View your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
    else:
        print("\n⚠️ Export to Google Sheets failed, but local backups were saved")

if __name__ == "__main__":
    main()