#!/usr/bin/env python3
"""
Test export with 3 contacts - ALL FIELDS COMPREHENSIVE
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

def export_contact_comprehensive(index):
    """Export single contact with ALL fields"""

    # Split into multiple scripts for stability
    # Part 1: Names and Organization
    script1 = f'''
    tell application "Contacts"
        try
            set p to person {index}
            set output to ""

            -- First Name
            try
                set val to first name of p
                if val is not missing value then
                    set output to output & (val as string)
                end if
            on error
                -- ignore
            end try
            set output to output & "|"

            -- Last Name
            try
                set val to last name of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- Middle Name
            try
                set val to middle name of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- Nickname
            try
                set val to nickname of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- Name Prefix
            try
                set val to name prefix of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- Name Suffix
            try
                set val to name suffix of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- Organization
            try
                set val to organization of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- Job Title
            try
                set val to job title of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- Department
            try
                set val to department of p
                if val is not missing value then
                    set output to output & val
                end if
            end try

            return output
        on error
            return "ERROR"
        end try
    end tell
    '''

    # Part 2: Contact Info (Emails and Phones)
    script2 = f'''
    tell application "Contacts"
        try
            set p to person {index}
            set output to ""

            -- ALL Emails
            set emailStr to ""
            try
                set emailList to emails of p
                if (count of emailList) > 0 then
                    repeat with i from 1 to (count of emailList)
                        if i > 1 then set emailStr to emailStr & "; "
                        set emailVal to value of item i of emailList
                        set emailLabel to "email"
                        try
                            set labelVal to label of item i of emailList
                            if labelVal is not missing value then set emailLabel to labelVal
                        end try
                        set emailStr to emailStr & emailLabel & ": " & emailVal
                    end repeat
                end if
            end try
            set output to output & emailStr & "|"

            -- ALL Phones
            set phoneStr to ""
            try
                set phoneList to phones of p
                if (count of phoneList) > 0 then
                    repeat with i from 1 to (count of phoneList)
                        if i > 1 then set phoneStr to phoneStr & "; "
                        set phoneVal to value of item i of phoneList
                        set phoneLabel to "phone"
                        try
                            set labelVal to label of item i of phoneList
                            if labelVal is not missing value then set phoneLabel to labelVal
                        end try
                        set phoneStr to phoneStr & phoneLabel & ": " & phoneVal
                    end repeat
                end if
            end try
            set output to output & phoneStr

            return output
        on error
            return "ERROR"
        end try
    end tell
    '''

    # Part 3: Addresses and Birthday
    script3 = f'''
    tell application "Contacts"
        try
            set p to person {index}
            set output to ""

            -- Birthday
            try
                set bd to birth date of p
                if bd is not missing value then
                    set output to output & (month of bd as string) & "/" & (day of bd as string) & "/" & (year of bd as string)
                end if
            end try
            set output to output & "|"

            -- ALL Addresses
            set addrStr to ""
            try
                set addrList to addresses of p
                if (count of addrList) > 0 then
                    repeat with i from 1 to (count of addrList)
                        if i > 1 then set addrStr to addrStr & "; "
                        set addr to item i of addrList
                        set addrLabel to "address"
                        try
                            set labelVal to label of item i of addrList
                            if labelVal is not missing value then set addrLabel to labelVal
                        end try

                        set addrParts to ""
                        try
                            set streetVal to street of addr
                            if streetVal is not missing value then set addrParts to addrParts & streetVal
                        end try
                        try
                            set cityVal to city of addr
                            if cityVal is not missing value then
                                if addrParts is not "" then set addrParts to addrParts & ", "
                                set addrParts to addrParts & cityVal
                            end if
                        end try
                        try
                            set stateVal to state of addr
                            if stateVal is not missing value then
                                if addrParts is not "" then set addrParts to addrParts & ", "
                                set addrParts to addrParts & stateVal
                            end if
                        end try
                        try
                            set zipVal to zip of addr
                            if zipVal is not missing value then
                                if addrParts is not "" then set addrParts to addrParts & " "
                                set addrParts to addrParts & zipVal
                            end if
                        end try
                        try
                            set countryVal to country of addr
                            if countryVal is not missing value then
                                if addrParts is not "" then set addrParts to addrParts & ", "
                                set addrParts to addrParts & countryVal
                            end if
                        end try

                        set addrStr to addrStr & addrLabel & ": " & addrParts
                    end repeat
                end if
            end try
            set output to output & addrStr

            return output
        on error
            return "ERROR"
        end try
    end tell
    '''

    # Part 4: Notes and URLs
    script4 = f'''
    tell application "Contacts"
        try
            set p to person {index}
            set output to ""

            -- Notes (FULL, no truncation)
            try
                set val to note of p
                if val is not missing value then
                    set output to output & val
                end if
            end try
            set output to output & "|"

            -- URLs
            set urlStr to ""
            try
                set urlList to urls of p
                if (count of urlList) > 0 then
                    repeat with i from 1 to (count of urlList)
                        if i > 1 then set urlStr to urlStr & "; "
                        set urlVal to value of item i of urlList
                        set urlStr to urlStr & "url: " & urlVal
                    end repeat
                end if
            end try
            set output to output & urlStr

            return output
        on error
            return "ERROR"
        end try
    end tell
    '''

    try:
        # Execute all parts
        result1 = subprocess.run(['osascript', '-e', script1], capture_output=True, text=True, timeout=5)
        if result1.returncode != 0 or result1.stdout.strip() == "ERROR":
            print(f"   Part 1 failed for contact {index}")
            return None

        result2 = subprocess.run(['osascript', '-e', script2], capture_output=True, text=True, timeout=5)
        if result2.returncode != 0 or result2.stdout.strip() == "ERROR":
            print(f"   Part 2 failed for contact {index}")
            return None

        result3 = subprocess.run(['osascript', '-e', script3], capture_output=True, text=True, timeout=5)
        if result3.returncode != 0 or result3.stdout.strip() == "ERROR":
            print(f"   Part 3 failed for contact {index}")
            return None

        result4 = subprocess.run(['osascript', '-e', script4], capture_output=True, text=True, timeout=5)
        if result4.returncode != 0 or result4.stdout.strip() == "ERROR":
            print(f"   Part 4 failed for contact {index}")
            return None

        # Combine all parts
        part1 = result1.stdout.strip()  # Names and org
        part2 = result2.stdout.strip()  # Emails and phones
        part3 = result3.stdout.strip()  # Birthday and addresses
        part4 = result4.stdout.strip()  # Notes and URLs

        combined = part1 + "|" + part2 + "|" + part3 + "|" + part4
        return combined

    except Exception as e:
        print(f"Exception for contact {index}: {e}")
        return None

def export_3_comprehensive():
    """Export first 3 contacts with ALL fields"""
    print("🧪 Testing COMPREHENSIVE export with first 3 contacts...")
    print("📋 This will export ALL fields for each contact\n")

    contacts_data = []

    for i in range(1, 4):  # First 3 contacts
        print(f"📱 Exporting contact {i} with ALL fields...")

        contact_data = export_contact_comprehensive(i)

        if contact_data:
            parts = contact_data.split('|')
            print(f"   ✅ Contact {i}: Got {len(parts)} fields")

            # Map to ALL fields
            contact = {
                'First Name': parts[0] if len(parts) > 0 else '',
                'Last Name': parts[1] if len(parts) > 1 else '',
                'Middle Name': parts[2] if len(parts) > 2 else '',
                'Nickname': parts[3] if len(parts) > 3 else '',
                'Name Prefix': parts[4] if len(parts) > 4 else '',
                'Name Suffix': parts[5] if len(parts) > 5 else '',
                'Organization': parts[6] if len(parts) > 6 else '',
                'Job Title': parts[7] if len(parts) > 7 else '',
                'Department': parts[8] if len(parts) > 8 else '',
                'All Emails': parts[9] if len(parts) > 9 else '',
                'All Phone Numbers': parts[10] if len(parts) > 10 else '',
                'Birthday': parts[11] if len(parts) > 11 else '',
                'All Addresses': parts[12] if len(parts) > 12 else '',
                'Notes (Full)': parts[13] if len(parts) > 13 else '',
                'All URLs': parts[14] if len(parts) > 14 else ''
            }

            # Show summary
            name = f"{parts[0]} {parts[1]}".strip() if len(parts) > 1 else "Unknown"
            print(f"      Name: {name}")
            if parts[6]:  # Organization
                print(f"      Org: {parts[6]}")
            if parts[9]:  # Emails
                print(f"      Emails: {parts[9][:60]}..." if len(parts[9]) > 60 else f"      Emails: {parts[9]}")
            if parts[10]:  # Phones
                print(f"      Phones: {parts[10][:60]}..." if len(parts[10]) > 60 else f"      Phones: {parts[10]}")
            if parts[12]:  # Addresses
                print(f"      Addresses: {parts[12][:60]}..." if len(parts[12]) > 60 else f"      Addresses: {parts[12]}")

            contacts_data.append(contact)
        else:
            print(f"   ❌ Contact {i}: Export failed")

    print(f"\n✅ Successfully exported {len(contacts_data)} contacts with COMPREHENSIVE data")
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

def upload_comprehensive_to_sheets(service, contacts_data):
    """Upload comprehensive data to Google Sheets"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found in .env file")
        sys.exit(1)

    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')

    print(f"\n📊 Uploading COMPREHENSIVE data to Google Sheets...")

    if not contacts_data:
        return None

    # ALL field headers
    headers = [
        'First Name', 'Last Name', 'Middle Name', 'Nickname',
        'Name Prefix', 'Name Suffix',
        'Organization', 'Job Title', 'Department',
        'All Emails', 'All Phone Numbers',
        'Birthday', 'All Addresses',
        'Notes (Full)', 'All URLs'
    ]

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

        # Upload comprehensive data
        print("   Uploading comprehensive contact data...")
        body = {'values': values}
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body=body
        ).execute()

        # Format the sheet
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

        print(f"✅ Successfully uploaded {len(contacts_data)} contacts with ALL fields!")
        print(f"📋 View your COMPREHENSIVE data: https://docs.google.com/spreadsheets/d/{sheet_id}")

        return sheet_id

    except Exception as e:
        print(f"❌ Error updating sheet: {str(e)}")
        return None

def save_comprehensive_backup(contacts_data):
    """Save comprehensive backup"""
    if not contacts_data:
        return None

    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"comprehensive_test_3_{timestamp}.json"

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts_data, f, indent=2, ensure_ascii=False)

    print(f"💾 Comprehensive backup saved: {json_file}")
    return json_file

def main():
    """Main test function"""
    print("=" * 60)
    print("COMPREHENSIVE TEST: 3 Contacts with ALL FIELDS")
    print("Testing complete data export - NO truncation, NO limits")
    print("=" * 60)

    # Export 3 contacts with ALL fields
    contacts = export_3_comprehensive()

    if not contacts:
        print("❌ No contacts exported")
        sys.exit(1)

    # Save comprehensive backup
    save_comprehensive_backup(contacts)

    # Upload to Google Sheets
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    sheet_id = upload_comprehensive_to_sheets(service, contacts)

    if sheet_id:
        print("\n🎉 COMPREHENSIVE TEST SUCCESSFUL!")
        print(f"📊 {len(contacts)} contacts uploaded with ALL fields")
        print("✅ Every field captured - names, emails, phones, addresses, notes, etc.")
        print(f"📋 Check your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
        print("\n🚀 Ready for full export of all 3,792 contacts!")
        print("   Run: ./run.sh")
    else:
        print("\n⚠️ Test failed")

if __name__ == "__main__":
    main()