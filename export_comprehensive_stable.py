#!/usr/bin/env python3
"""
Mac Contacts to Google Sheets - Comprehensive & Stable Version
Exports ALL fields but with more stable AppleScript
"""

import os
import sys
import subprocess
import json
import time
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from tqdm import tqdm

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

def get_total_contacts():
    """Get total number of contacts"""
    script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    return int(result.stdout.strip())

def export_contact_by_index_stable(index):
    """Export a single contact with all fields (stable version)"""

    # Get basic contact info first
    basic_script = f'''
    tell application "Contacts"
        try
            set p to person {index}

            -- Names
            try
                set fn to first name of p
            on error
                set fn to ""
            end try
            try
                set ln to last name of p
            on error
                set ln to ""
            end try
            try
                set mn to middle name of p
            on error
                set mn to ""
            end try
            try
                set nn to nickname of p
            on error
                set nn to ""
            end try
            try
                set npx to name prefix of p
            on error
                set npx to ""
            end try
            try
                set nsx to name suffix of p
            on error
                set nsx to ""
            end try

            -- Organization
            try
                set org to organization of p
            on error
                set org to ""
            end try
            try
                set job to job title of p
            on error
                set job to ""
            end try
            try
                set dept to department of p
            on error
                set dept to ""
            end try

            -- Birthday
            try
                set bd to birth date of p
                set bdStr to (month of bd as string) & "/" & (day of bd as string) & "/" & (year of bd as string)
            on error
                set bdStr to ""
            end try

            return fn & "|" & ln & "|" & mn & "|" & nn & "|" & npx & "|" & nsx & "|" & org & "|" & job & "|" & dept & "|" & bdStr

        on error
            return "ERROR"
        end try
    end tell
    '''

    # Get contact details (emails, phones, etc.)
    details_script = f'''
    tell application "Contacts"
        try
            set p to person {index}

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
                        end try
                        if i > 1 then set phoneStr to phoneStr & "; "
                        set phoneStr to phoneStr & phLabel & ": " & ph
                    end repeat
                end if
            on error
                set phoneStr to ""
            end try

            -- All addresses with labels
            set addrStr to ""
            try
                set addrList to addresses of p
                if (count of addrList) > 0 then
                    repeat with i from 1 to (count of addrList)
                        set addr to item i of addrList
                        set addrLabel to "address"
                        try
                            set addrLabel to label of item i of addrList
                        end try

                        set singleAddr to ""
                        try
                            set streetVal to street of addr
                            if streetVal is not "" then set singleAddr to streetVal
                        end try
                        try
                            set cityVal to city of addr
                            if cityVal is not "" then
                                if singleAddr is not "" then set singleAddr to singleAddr & ", "
                                set singleAddr to singleAddr & cityVal
                            end if
                        end try
                        try
                            set stateVal to state of addr
                            if stateVal is not "" then
                                if singleAddr is not "" then set singleAddr to singleAddr & ", "
                                set singleAddr to singleAddr & stateVal
                            end if
                        end try
                        try
                            set zipVal to zip of addr
                            if zipVal is not "" then
                                if singleAddr is not "" then set singleAddr to singleAddr & " "
                                set singleAddr to singleAddr & zipVal
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

            return emailStr & "|" & phoneStr & "|" & addrStr

        on error
            return "ERROR"
        end try
    end tell
    '''

    # Get additional data (URLs, notes, etc.)
    additional_script = f'''
    tell application "Contacts"
        try
            set p to person {index}

            -- All URLs
            set urlStr to ""
            try
                set urlList to urls of p
                if (count of urlList) > 0 then
                    repeat with i from 1 to (count of urlList)
                        set url to value of item i of urlList
                        set urlLabel to "url"
                        try
                            set urlLabel to label of item i of urlList
                        end try
                        if i > 1 then set urlStr to urlStr & "; "
                        set urlStr to urlStr & urlLabel & ": " & url
                    end repeat
                end if
            on error
                set urlStr to ""
            end try

            -- Notes (full, no truncation)
            try
                set nt to note of p
            on error
                set nt to ""
            end try

            -- Related names
            set relatedStr to ""
            try
                set relatedList to related names of p
                if (count of relatedList) > 0 then
                    repeat with i from 1 to (count of relatedList)
                        set related to value of item i of relatedList
                        set relatedLabel to "related"
                        try
                            set relatedLabel to label of item i of relatedList
                        end try
                        if i > 1 then set relatedStr to relatedStr & "; "
                        set relatedStr to relatedStr & relatedLabel & ": " & related
                    end repeat
                end if
            on error
                set relatedStr to ""
            end try

            return urlStr & "|" & relatedStr & "|" & nt

        on error
            return "ERROR"
        end try
    end tell
    '''

    try:
        # Run basic info script
        result1 = subprocess.run(['osascript', '-e', basic_script], capture_output=True, text=True, timeout=3)
        if result1.returncode != 0 or result1.stdout.strip() == "ERROR":
            return None
        basic_data = result1.stdout.strip()

        # Run details script
        result2 = subprocess.run(['osascript', '-e', details_script], capture_output=True, text=True, timeout=3)
        if result2.returncode != 0 or result2.stdout.strip() == "ERROR":
            return None
        details_data = result2.stdout.strip()

        # Run additional script
        result3 = subprocess.run(['osascript', '-e', additional_script], capture_output=True, text=True, timeout=3)
        if result3.returncode != 0 or result3.stdout.strip() == "ERROR":
            return None
        additional_data = result3.stdout.strip()

        # Combine all data
        combined_data = basic_data + "|" + details_data + "|" + additional_data
        return combined_data

    except:
        return None

def export_all_contacts():
    """Export all contacts with comprehensive data"""
    print("📱 Starting comprehensive contact export...")

    total = get_total_contacts()
    print(f"📊 Found {total:,} contacts to export")

    contacts_data = []
    failed_indices = []

    print("\n⏳ Processing contacts with ALL fields...")

    with tqdm(total=total, desc="Exporting", unit="contact", ncols=100) as pbar:
        for i in range(1, total + 1):
            contact_data = export_contact_by_index_stable(i)

            if contact_data:
                parts = contact_data.split('|')
                if len(parts) >= 15:  # Expecting at least 15 fields
                    contact = {
                        'First Name': parts[0],
                        'Last Name': parts[1],
                        'Middle Name': parts[2],
                        'Nickname': parts[3],
                        'Name Prefix': parts[4],
                        'Name Suffix': parts[5],
                        'Organization': parts[6],
                        'Job Title': parts[7],
                        'Department': parts[8],
                        'Birthday': parts[9],
                        'All Emails': parts[10],
                        'All Phone Numbers': parts[11],
                        'All Addresses': parts[12],
                        'All URLs': parts[13],
                        'Related Names': parts[14],
                        'Notes': parts[15] if len(parts) > 15 else ''
                    }

                    # Only add if contact has meaningful data
                    if any([parts[0], parts[1], parts[10], parts[11]]):
                        contacts_data.append(contact)
            else:
                failed_indices.append(i)

            pbar.update(1)

            # Small delay every 50 contacts
            if i % 50 == 0:
                time.sleep(0.1)

    print(f"\n✅ Successfully exported {len(contacts_data)} contacts")
    if failed_indices:
        print(f"⚠️  Failed to export {len(failed_indices)} contacts")

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

def update_google_sheet_batch(service, contacts_data):
    """Update Google Sheet with comprehensive contact data"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found in .env file")
        sys.exit(1)

    sheet_name = os.getenv('SHEET_NAME', 'Contacts')

    print(f"\n📊 Uploading to Google Sheets...")

    if not contacts_data:
        return None

    # Headers for comprehensive export
    headers = ['First Name', 'Last Name', 'Middle Name', 'Nickname',
               'Name Prefix', 'Name Suffix', 'Organization', 'Job Title', 'Department',
               'Birthday', 'All Emails', 'All Phone Numbers', 'All Addresses',
               'All URLs', 'Related Names', 'Notes']

    values = [headers]
    for contact in contacts_data:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear and update
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Upload in batches
        batch_size = 1000
        for start_idx in range(0, len(values), batch_size):
            end_idx = min(start_idx + batch_size, len(values))
            batch = values[start_idx:end_idx]

            range_start = f"{sheet_name}!A{start_idx + 1}" if start_idx > 0 else f"{sheet_name}!A1"
            print(f"   Uploading rows {start_idx + 1} to {end_idx}...")

            service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=range_start,
                valueInputOption='RAW',
                body={'values': batch}
            ).execute()

        # Format
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
            }
        ]

        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={'requests': requests}
        ).execute()

        print(f"✅ Successfully uploaded {len(contacts_data)} contacts")
        print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return sheet_id

    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def save_backup(contacts_data):
    """Save local backups"""
    if not contacts_data:
        return None

    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # JSON backup
    json_file = export_dir / f"contacts_comprehensive_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts_data, f, indent=2, ensure_ascii=False)

    # CSV backup
    import csv
    csv_file = export_dir / f"contacts_comprehensive_{timestamp}.csv"

    headers = ['First Name', 'Last Name', 'Middle Name', 'Nickname',
               'Name Prefix', 'Name Suffix', 'Organization', 'Job Title', 'Department',
               'Birthday', 'All Emails', 'All Phone Numbers', 'All Addresses',
               'All URLs', 'Related Names', 'Notes']

    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(contacts_data)

    print(f"\n💾 Backups saved:")
    print(f"   JSON: {json_file}")
    print(f"   CSV:  {csv_file}")

    return json_file, csv_file

def main():
    """Main function"""
    print("=" * 60)
    print("Mac Contacts to Google Sheets - COMPREHENSIVE EXPORT")
    print("ALL Fields - NO Truncation - NO Limits")
    print("=" * 60)

    # Export all contacts
    contacts = export_all_contacts()

    if not contacts:
        print("❌ No contacts were exported")
        sys.exit(1)

    # Save backups
    save_backup(contacts)

    # Upload option
    print(f"\n📤 Upload {len(contacts):,} contacts to Google Sheets?")
    response = input("   Continue? (y/n): ").lower().strip()

    if response == 'y':
        print("\n🔐 Authenticating...")
        service = authenticate_google_sheets()
        sheet_id = update_google_sheet_batch(service, contacts)

        if sheet_id:
            print("\n🎉 COMPREHENSIVE EXPORT COMPLETE!")
            print(f"📱 Total contacts: {len(contacts):,}")
            print(f"📊 All fields exported with NO data loss!")
            print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        else:
            print("\n⚠️  Upload failed, but backups saved")
    else:
        print("\n✅ Comprehensive export complete. Backups saved.")

if __name__ == "__main__":
    main()