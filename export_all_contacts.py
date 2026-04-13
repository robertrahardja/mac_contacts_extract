#!/usr/bin/env python3
"""
Mac Contacts to Google Sheets - No Timeout Version
Processes contacts one by one to avoid timeouts
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

def export_contact_by_index(index):
    """Export a single contact by index"""
    applescript = f'''
    tell application "Contacts"
        try
            set p to person {index}

            -- Get names
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

            -- Get organization
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

            -- Get ALL emails
            set emailStr to ""
            try
                set emailList to emails of p
                set emailCount to count of emailList
                if emailCount > 0 then
                    repeat with i from 1 to emailCount
                        set em to value of item i of emailList
                        try
                            set emLabel to label of item i of emailList
                        on error
                            set emLabel to "other"
                        end try
                        if i > 1 then set emailStr to emailStr & "; "
                        set emailStr to emailStr & emLabel & ": " & em
                    end repeat
                end if
            on error
                set emailStr to ""
            end try

            -- Get ALL phone numbers
            set phoneStr to ""
            try
                set phoneList to phones of p
                set phoneCount to count of phoneList
                if phoneCount > 0 then
                    repeat with i from 1 to phoneCount
                        set ph to value of item i of phoneList
                        try
                            set phLabel to label of item i of phoneList
                        on error
                            set phLabel to "other"
                        end try
                        if i > 1 then set phoneStr to phoneStr & "; "
                        set phoneStr to phoneStr & phLabel & ": " & ph
                    end repeat
                end if
            on error
                set phoneStr to ""
            end try

            -- Get birthday
            try
                set bd to birth date of p
                set bdStr to (month of bd as string) & "/" & (day of bd as string) & "/" & (year of bd as string)
            on error
                set bdStr to ""
            end try

            -- Get FULL note (no truncation)
            try
                set nt to note of p
            on error
                set nt to ""
            end try

            -- Get ALL addresses
            set addrStr to ""
            try
                set addrList to addresses of p
                set addrCount to count of addrList
                if addrCount > 0 then
                    repeat with i from 1 to addrCount
                        set addr to item i of addrList
                        try
                            set addrLabel to label of item i of addrList
                        on error
                            set addrLabel to "other"
                        end try

                        set singleAddr to ""
                        try
                            set streetVal to street of addr
                            if streetVal is not "" then set singleAddr to singleAddr & streetVal
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
                        try
                            set countryVal to country of addr
                            if countryVal is not "" then
                                if singleAddr is not "" then set singleAddr to singleAddr & ", "
                                set singleAddr to singleAddr & countryVal
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

            -- Get ALL URLs
            set urlStr to ""
            try
                set urlList to urls of p
                set urlCount to count of urlList
                if urlCount > 0 then
                    repeat with i from 1 to urlCount
                        set url to value of item i of urlList
                        try
                            set urlLabel to label of item i of urlList
                        on error
                            set urlLabel to "other"
                        end try
                        if i > 1 then set urlStr to urlStr & "; "
                        set urlStr to urlStr & urlLabel & ": " & url
                    end repeat
                end if
            on error
                set urlStr to ""
            end try

            -- Get ALL social profiles
            set socialStr to ""
            try
                set socialList to social profiles of p
                set socialCount to count of socialList
                if socialCount > 0 then
                    repeat with i from 1 to socialCount
                        set social to item i of socialList
                        try
                            set socialService to service of social
                        on error
                            set socialService to "unknown"
                        end try
                        try
                            set socialUsername to username of social
                        on error
                            set socialUsername to ""
                        end try
                        try
                            set socialUrl to url of social
                        on error
                            set socialUrl to ""
                        end try

                        if socialUsername is not "" or socialUrl is not "" then
                            if i > 1 then set socialStr to socialStr & "; "
                            set socialStr to socialStr & socialService & ": " & socialUsername
                            if socialUrl is not "" then set socialStr to socialStr & " (" & socialUrl & ")"
                        end if
                    end repeat
                end if
            on error
                set socialStr to ""
            end try

            -- Get ALL instant message accounts
            set imStr to ""
            try
                set imList to instant message addresses of p
                set imCount to count of imList
                if imCount > 0 then
                    repeat with i from 1 to imCount
                        set im to item i of imList
                        try
                            set imService to service of im
                        on error
                            set imService to "unknown"
                        end try
                        try
                            set imUsername to username of im
                        on error
                            set imUsername to ""
                        end try

                        if imUsername is not "" then
                            if i > 1 then set imStr to imStr & "; "
                            set imStr to imStr & imService & ": " & imUsername
                        end if
                    end repeat
                end if
            on error
                set imStr to ""
            end try

            -- Get ALL related names
            set relatedStr to ""
            try
                set relatedList to related names of p
                set relatedCount to count of relatedList
                if relatedCount > 0 then
                    repeat with i from 1 to relatedCount
                        set related to value of item i of relatedList
                        try
                            set relatedLabel to label of item i of relatedList
                        on error
                            set relatedLabel to "other"
                        end try
                        if i > 1 then set relatedStr to relatedStr & "; "
                        set relatedStr to relatedStr & relatedLabel & ": " & related
                    end repeat
                end if
            on error
                set relatedStr to ""
            end try

            -- Get department
            try
                set dept to department of p
            on error
                set dept to ""
            end try

            -- Get name suffix
            try
                set suffix to name suffix of p
            on error
                set suffix to ""
            end try

            -- Get name prefix
            try
                set prefix to name prefix of p
            on error
                set prefix to ""
            end try

            -- Get phonetic names
            try
                set phoneticFirst to phonetic first name of p
            on error
                set phoneticFirst to ""
            end try

            try
                set phoneticLast to phonetic last name of p
            on error
                set phoneticLast to ""
            end try

            try
                set phoneticMiddle to phonetic middle name of p
            on error
                set phoneticMiddle to ""
            end try

            -- Return all fields pipe-separated
            return fn & "|" & ln & "|" & mn & "|" & nn & "|" & prefix & "|" & suffix & "|" & phoneticFirst & "|" & phoneticMiddle & "|" & phoneticLast & "|" & org & "|" & job & "|" & dept & "|" & emailStr & "|" & phoneStr & "|" & bdStr & "|" & addrStr & "|" & urlStr & "|" & socialStr & "|" & imStr & "|" & relatedStr & "|" & nt

        on error
            return "ERROR"
        end try
    end tell
    '''

    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            timeout=5  # 5 second timeout per contact for comprehensive data
        )

        if result.returncode == 0 and result.stdout.strip() != "ERROR":
            return result.stdout.strip()
        return None
    except:
        return None

def minimum(a, b):
    """Helper to get minimum of two numbers"""
    return a if a < b else b

def export_all_contacts():
    """Export all contacts with progress tracking"""
    print("📱 Starting contact export...")

    # Get total count
    total = get_total_contacts()
    print(f"📊 Found {total:,} contacts to export")

    contacts_data = []
    failed_indices = []

    # Process contacts with progress bar
    print("\n⏳ Processing contacts (this may take a few minutes)...")

    with tqdm(total=total, desc="Exporting", unit="contact", ncols=100) as pbar:
        for i in range(1, min(total + 1, total + 1)):  # Process ALL contacts
            # Export single contact
            contact_data = export_contact_by_index(i)

            if contact_data:
                parts = contact_data.split('|')
                if len(parts) >= 21:
                    contact = {
                        'First Name': parts[0],
                        'Last Name': parts[1],
                        'Middle Name': parts[2],
                        'Nickname': parts[3],
                        'Name Prefix': parts[4],
                        'Name Suffix': parts[5],
                        'Phonetic First Name': parts[6],
                        'Phonetic Middle Name': parts[7],
                        'Phonetic Last Name': parts[8],
                        'Organization': parts[9],
                        'Job Title': parts[10],
                        'Department': parts[11],
                        'All Emails': parts[12],
                        'All Phone Numbers': parts[13],
                        'Birthday': parts[14],
                        'All Addresses': parts[15],
                        'All URLs': parts[16],
                        'Social Profiles': parts[17],
                        'Instant Messages': parts[18],
                        'Related Names': parts[19],
                        'Notes': parts[20] if len(parts) > 20 else ''
                    }

                    # Only add if contact has some meaningful data
                    if any([parts[0], parts[1], parts[12], parts[13]]):  # Has name or contact info
                        contacts_data.append(contact)
            else:
                failed_indices.append(i)

            pbar.update(1)

            # Small delay every 100 contacts to avoid overwhelming the system
            if i % 100 == 0:
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
    """Update Google Sheet with batched data"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found in .env file")
        sys.exit(1)

    sheet_name = os.getenv('SHEET_NAME', 'Contacts')

    print(f"\n📊 Uploading to Google Sheets...")

    if not contacts_data:
        print("No contacts to upload")
        return None

    # Prepare headers and values - ALL CONTACT FIELDS
    headers = ['First Name', 'Last Name', 'Middle Name', 'Nickname',
               'Name Prefix', 'Name Suffix', 'Phonetic First Name', 'Phonetic Middle Name', 'Phonetic Last Name',
               'Organization', 'Job Title', 'Department',
               'All Emails', 'All Phone Numbers', 'Birthday', 'All Addresses',
               'All URLs', 'Social Profiles', 'Instant Messages', 'Related Names', 'Notes']

    values = [headers]
    for contact in contacts_data:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear existing content
        print("   Clearing existing data...")
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Upload in batches of 1000 rows to avoid API limits
        batch_size = 1000
        total_rows = len(values)

        for start_idx in range(0, total_rows, batch_size):
            end_idx = min(start_idx + batch_size, total_rows)
            batch = values[start_idx:end_idx]

            if start_idx == 0:
                # First batch includes headers
                range_start = f"{sheet_name}!A1"
            else:
                # Subsequent batches start after headers
                range_start = f"{sheet_name}!A{start_idx + 1}"

            print(f"   Uploading rows {start_idx + 1} to {end_idx}...")

            body = {'values': batch}
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=range_start,
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
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': 0,
                        'endRowIndex': 1
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {'bold': True}
                        }
                    },
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

        print(f"✅ Successfully uploaded {len(contacts_data)} contacts")
        print(f"📋 View your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")

        return sheet_id

    except Exception as e:
        print(f"❌ Error updating sheet: {str(e)}")
        return None

def save_backup(contacts_data):
    """Save local backup of contacts"""
    if not contacts_data:
        return None

    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Save as JSON
    json_file = export_dir / f"contacts_full_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts_data, f, indent=2, ensure_ascii=False)

    # Save as CSV
    import csv
    csv_file = export_dir / f"contacts_full_{timestamp}.csv"

    if contacts_data:
        headers = ['First Name', 'Last Name', 'Middle Name', 'Nickname',
                   'Name Prefix', 'Name Suffix', 'Phonetic First Name', 'Phonetic Middle Name', 'Phonetic Last Name',
                   'Organization', 'Job Title', 'Department',
                   'All Emails', 'All Phone Numbers', 'Birthday', 'All Addresses',
                   'All URLs', 'Social Profiles', 'Instant Messages', 'Related Names', 'Notes']

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
    print("Mac Contacts to Google Sheets - Complete Export")
    print("No Timeout Version - Processes All Contacts")
    print("=" * 60)

    # Export all contacts
    contacts = export_all_contacts()

    if not contacts:
        print("❌ No contacts were exported")
        sys.exit(1)

    # Save local backups
    save_backup(contacts)

    # Ask if user wants to upload to Google Sheets
    print("\n📤 Upload to Google Sheets?")
    print("   This will replace all data in your sheet.")
    response = input("   Continue? (y/n): ").lower().strip()

    if response == 'y':
        # Authenticate with Google
        print("\n🔐 Authenticating with Google...")
        service = authenticate_google_sheets()

        # Update Google Sheet
        sheet_id = update_google_sheet_batch(service, contacts)

        if sheet_id:
            print("\n✨ Export complete!")
            print(f"📱 Total contacts exported: {len(contacts):,}")
            print(f"📊 View your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
        else:
            print("\n⚠️  Export to Google Sheets failed, but local backups were saved")
    else:
        print("\n✅ Local export complete. Backups saved in exports/ folder.")

if __name__ == "__main__":
    main()