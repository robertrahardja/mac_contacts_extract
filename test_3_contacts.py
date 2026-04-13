#!/usr/bin/env python3
"""
Test export with just 3 contacts to verify Google Sheets upload
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

def export_contact_fixed(index):
    """Export single contact with fixed AppleScript (handles missing values)"""

    applescript = f'''
    tell application "Contacts"
        try
            set p to person {index}

            -- Names (handle missing values)
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
                set mn to middle name of p
                if mn is missing value then set mn to ""
            on error
                set mn to ""
            end try

            try
                set nn to nickname of p
                if nn is missing value then set nn to ""
            on error
                set nn to ""
            end try

            -- Organization
            try
                set org to organization of p
                if org is missing value then set org to ""
            on error
                set org to ""
            end try

            try
                set job to job title of p
                if job is missing value then set job to ""
            on error
                set job to ""
            end try

            -- Birthday
            try
                set bd to birth date of p
                if bd is not missing value then
                    set bdStr to (month of bd as string) & "/" & (day of bd as string) & "/" & (year of bd as string)
                else
                    set bdStr to ""
                end if
            on error
                set bdStr to ""
            end try

            -- All emails with labels
            set emailStr to ""
            try
                set emailList to emails of p
                if (count of emailList) > 0 then
                    repeat with i from 1 to (count of emailList)
                        set em to value of item i of emailList
                        set emLabel to "email"
                        try
                            set emLabel to label of item i of emailList
                            if emLabel is missing value then set emLabel to "email"
                        end try
                        if i > 1 then set emailStr to emailStr & "; "
                        set emailStr to emailStr & emLabel & ": " & em
                    end repeat
                end if
            on error
                set emailStr to ""
            end try

            -- All phones with labels
            set phoneStr to ""
            try
                set phoneList to phones of p
                if (count of phoneList) > 0 then
                    repeat with i from 1 to (count of phoneList)
                        set ph to value of item i of phoneList
                        set phLabel to "phone"
                        try
                            set phLabel to label of item i of phoneList
                            if phLabel is missing value then set phLabel to "phone"
                        end try
                        if i > 1 then set phoneStr to phoneStr & "; "
                        set phoneStr to phoneStr & phLabel & ": " & ph
                    end repeat
                end if
            on error
                set phoneStr to ""
            end try

            -- All addresses
            set addrStr to ""
            try
                set addrList to addresses of p
                if (count of addrList) > 0 then
                    repeat with i from 1 to (count of addrList)
                        set addr to item i of addrList
                        set addrLabel to "address"
                        try
                            set addrLabel to label of item i of addrList
                            if addrLabel is missing value then set addrLabel to "address"
                        end try

                        set singleAddr to ""
                        try
                            set streetVal to street of addr
                            if streetVal is not missing value then set singleAddr to streetVal as string
                        end try
                        try
                            set cityVal to city of addr
                            if cityVal is not missing value then
                                if singleAddr is not "" then set singleAddr to singleAddr & ", "
                                set singleAddr to singleAddr & (cityVal as string)
                            end if
                        end try
                        try
                            set stateVal to state of addr
                            if stateVal is not missing value then
                                if singleAddr is not "" then set singleAddr to singleAddr & ", "
                                set singleAddr to singleAddr & (stateVal as string)
                            end if
                        end try
                        try
                            set zipVal to zip of addr
                            if zipVal is not missing value then
                                if singleAddr is not "" then set singleAddr to singleAddr & " "
                                set singleAddr to singleAddr & (zipVal as string)
                            end if
                        end try

                        if singleAddr is not "" then
                            if i > 1 then set addrStr to addrStr & "; "
                            set addrStr to addrStr & addrLabel & ": " & singleAddr
                        end if
                    end repeat
                end if
            on error
                set addrStr to ""
            end try

            -- Notes (full)
            try
                set nt to note of p
                if nt is missing value then set nt to ""
            on error
                set nt to ""
            end try

            -- URLs
            set urlStr to ""
            try
                set urlList to urls of p
                if (count of urlList) > 0 then
                    repeat with i from 1 to (count of urlList)
                        set url to value of item i of urlList
                        set urlLabel to "url"
                        try
                            set urlLabel to label of item i of urlList
                            if urlLabel is missing value then set urlLabel to "url"
                        end try
                        if i > 1 then set urlStr to urlStr & "; "
                        set urlStr to urlStr & urlLabel & ": " & url
                    end repeat
                end if
            on error
                set urlStr to ""
            end try

            -- Return all data pipe-separated
            return fn & "|" & ln & "|" & mn & "|" & nn & "|" & org & "|" & job & "|" & bdStr & "|" & emailStr & "|" & phoneStr & "|" & addrStr & "|" & urlStr & "|" & nt

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
            timeout=10
        )

        if result.returncode == 0 and not result.stdout.strip().startswith("ERROR"):
            return result.stdout.strip()
        else:
            print(f"AppleScript error for contact {index}: {result.stderr}")
            return None
    except Exception as e:
        print(f"Exception for contact {index}: {e}")
        return None

def export_3_contacts():
    """Export first 3 contacts for testing"""
    print("🧪 Testing with first 3 contacts...")

    contacts_data = []

    for i in range(1, 4):  # Contacts 1, 2, 3
        print(f"   Exporting contact {i}...")

        contact_data = export_contact_fixed(i)

        if contact_data:
            parts = contact_data.split('|')
            print(f"   ✅ Contact {i}: Got {len(parts)} fields")

            if len(parts) >= 12:
                contact = {
                    'First Name': parts[0],
                    'Last Name': parts[1],
                    'Middle Name': parts[2],
                    'Nickname': parts[3],
                    'Organization': parts[4],
                    'Job Title': parts[5],
                    'Birthday': parts[6],
                    'All Emails': parts[7],
                    'All Phone Numbers': parts[8],
                    'All Addresses': parts[9],
                    'All URLs': parts[10],
                    'Notes': parts[11] if len(parts) > 11 else ''
                }

                # Show preview
                print(f"      Name: {parts[0]} {parts[1]}")
                print(f"      Org: {parts[4]}")
                print(f"      Emails: {parts[7][:50]}..." if len(parts[7]) > 50 else f"      Emails: {parts[7]}")

                contacts_data.append(contact)
            else:
                print(f"   ⚠️ Contact {i}: Insufficient fields ({len(parts)})")
        else:
            print(f"   ❌ Contact {i}: Export failed")

    print(f"\n✅ Successfully exported {len(contacts_data)} test contacts")
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
            if not os.path.exists('credentials.json'):
                print("❌ credentials.json not found!")
                sys.exit(1)

            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)

    return build('sheets', 'v4', credentials=creds)

def update_google_sheet_test(service, contacts_data):
    """Update Google Sheet with test data"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found in .env file")
        sys.exit(1)

    sheet_name = os.getenv('SHEET_NAME', 'Contacts')

    print(f"\n📊 Uploading {len(contacts_data)} test contacts to Google Sheets...")

    if not contacts_data:
        return None

    # Headers
    headers = ['First Name', 'Last Name', 'Middle Name', 'Nickname',
               'Organization', 'Job Title', 'Birthday',
               'All Emails', 'All Phone Numbers', 'All Addresses',
               'All URLs', 'Notes']

    values = [headers]
    for contact in contacts_data:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear sheet first
        print("   Clearing existing data...")
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Upload data
        print("   Uploading test data...")
        body = {'values': values}
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body=body
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

        print(f"✅ Successfully uploaded {len(contacts_data)} test contacts!")
        print(f"📋 View your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")

        return sheet_id

    except Exception as e:
        print(f"❌ Error updating sheet: {str(e)}")
        return None

def save_test_backup(contacts_data):
    """Save test backup"""
    if not contacts_data:
        return None

    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"test_3_contacts_{timestamp}.json"

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts_data, f, indent=2, ensure_ascii=False)

    print(f"💾 Test backup saved: {json_file}")
    return json_file

def main():
    """Main test function"""
    print("=" * 60)
    print("TEST: Export 3 Contacts to Google Sheets")
    print("Verify everything works before full export")
    print("=" * 60)

    # Export 3 test contacts
    contacts = export_3_contacts()

    if not contacts:
        print("❌ No test contacts exported")
        sys.exit(1)

    # Save test backup
    save_test_backup(contacts)

    # Upload to Google Sheets
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    sheet_id = update_google_sheet_test(service, contacts)

    if sheet_id:
        print("\n🎉 TEST SUCCESSFUL!")
        print(f"📊 {len(contacts)} contacts uploaded with comprehensive data")
        print(f"📋 Check your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
        print("\n✅ Ready for full export of all 3,792 contacts!")
        print("   Run: ./run.sh")
    else:
        print("\n⚠️ Test failed")

if __name__ == "__main__":
    main()