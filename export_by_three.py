#!/usr/bin/env python3
"""
Export ALL contacts - 3 at a time for reliability
"""

import os
import sys
import subprocess
import json
import time
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_three_contacts(start_idx):
    """Get 3 contacts at once - much faster than one by one"""
    script = f'''
    tell application "Contacts"
        set output to ""
        repeat with i from {start_idx} to {min(start_idx + 2, start_idx + 100)}
            try
                set p to person i
                set row to ""

                -- First Name
                try
                    set row to row & (first name of p)
                end try
                set row to row & "|"

                -- Last Name
                try
                    set row to row & (last name of p)
                end try
                set row to row & "|"

                -- Organization
                try
                    set row to row & (organization of p)
                end try
                set row to row & "|"

                -- First Email
                try
                    set emailList to emails of p
                    if (count of emailList) > 0 then
                        set row to row & (value of item 1 of emailList)
                    end if
                end try
                set row to row & "|"

                -- First Phone
                try
                    set phoneList to phones of p
                    if (count of phoneList) > 0 then
                        set row to row & (value of item 1 of phoneList)
                    end if
                end try

                if row is not "||||" then
                    if output is not "" then set output to output & "\\n"
                    set output to output & row
                end if
            on error
                -- Skip this contact
            end try
        end repeat
        return output
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script],
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            contacts = []
            lines = result.stdout.strip().split('\n')
            for line in lines:
                if line and line != '||||':
                    parts = line.split('|')
                    if len(parts) >= 5:
                        contact = {
                            'First Name': parts[0],
                            'Last Name': parts[1],
                            'Organization': parts[2],
                            'Primary Email': parts[3],
                            'Primary Phone': parts[4] if len(parts) > 4 else ''
                        }
                        if any(contact.values()):
                            contacts.append(contact)
            return contacts
    except:
        pass
    return []

def export_all_by_three():
    """Export all contacts, 3 at a time"""
    print("📦 Exporting ALL contacts (3 at a time for speed)...")

    # Get total count
    count_script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', count_script], capture_output=True, text=True)
    total = int(result.stdout.strip())
    print(f"📊 Total contacts to export: {total:,}\n")

    # Check for existing progress
    progress_file = Path('exports/three_at_time_progress.json')
    if progress_file.exists():
        with open(progress_file, 'r') as f:
            data = json.load(f)
            all_contacts = data.get('contacts', [])
            last_idx = data.get('last_index', 0)
            print(f"📂 Resuming from contact {last_idx}...")
            print(f"   Already exported: {len(all_contacts)}\n")
    else:
        all_contacts = []
        last_idx = 0
        progress_file.parent.mkdir(exist_ok=True)

    # Process 3 at a time
    batch_count = 0
    for i in range(last_idx + 1, total + 1, 3):
        batch_contacts = get_three_contacts(i)
        all_contacts.extend(batch_contacts)
        batch_count += 1

        # Progress update
        current = min(i + 2, total)
        percent = (current / total) * 100
        print(f"✅ Batch {batch_count}: Contacts {i}-{min(i+2, total)} ({percent:.1f}%) - Total found: {len(all_contacts)}")

        # Save progress every 30 contacts (10 batches)
        if batch_count % 10 == 0:
            with open(progress_file, 'w') as f:
                json.dump({'contacts': all_contacts, 'last_index': current}, f)
            print(f"   💾 Progress saved: {len(all_contacts)} contacts")

        # Brief pause every 100 batches to avoid overwhelming
        if batch_count % 100 == 0:
            print("   ⏸️ Brief pause...")
            time.sleep(1)

    # Final save
    with open(progress_file, 'w') as f:
        json.dump({'contacts': all_contacts, 'last_index': total}, f)

    # Save final backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_file = progress_file.parent / f"all_contacts_final_{timestamp}.json"
    with open(final_file, 'w') as f:
        json.dump(all_contacts, f, indent=2)

    print(f"\n🎉 EXPORT COMPLETE!")
    print(f"✅ Total contacts exported: {len(all_contacts)}")
    print(f"💾 Final backup saved: {final_file}")

    # Clean up progress file
    if progress_file.exists():
        progress_file.unlink()

    return all_contacts

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
    """Upload to Google Sheets"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')

    if not contacts:
        print("❌ No contacts to upload")
        return False

    print(f"\n📊 Uploading {len(contacts)} contacts to Google Sheets...")

    headers = ['First Name', 'Last Name', 'Organization', 'Primary Email', 'Primary Phone']
    values = [headers]

    for contact in contacts:
        row = [contact.get(h, '') for h in headers]
        values.append(row)

    try:
        # Clear sheet
        print("   Clearing sheet...")
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Upload in chunks to avoid size limits
        chunk_size = 5000
        for start in range(0, len(values), chunk_size):
            end = min(start + chunk_size, len(values))
            chunk = values[start:end]

            if start == 0:
                range_start = 'A1'
            else:
                range_start = f'A{start + 1}'

            print(f"   Uploading rows {start + 1} to {end}...")
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=f"{sheet_name}!{range_start}",
                valueInputOption='RAW',
                body={'values': chunk}
            ).execute()

        # Format headers
        print("   Formatting...")
        requests = [
            {
                'repeatCell': {
                    'range': {'sheetId': 0, 'startRowIndex': 0, 'endRowIndex': 1},
                    'cell': {'userEnteredFormat': {'textFormat': {'bold': True}}},
                    'fields': 'userEnteredFormat.textFormat.bold'
                }
            },
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': 0,
                        'gridProperties': {'frozenRowCount': 1}
                    },
                    'fields': 'gridProperties.frozenRowCount'
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

def main():
    print("="*60)
    print("EXPORT ALL 3,792 CONTACTS")
    print("Processing 3 at a time for optimal speed")
    print("="*60)

    start_time = time.time()

    # Export all contacts
    contacts = export_all_by_three()

    if not contacts:
        print("❌ No contacts exported")
        sys.exit(1)

    # Upload to Google Sheets
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    if upload_to_sheets(service, contacts):
        elapsed = time.time() - start_time
        print(f"\n🎉 COMPLETE!")
        print(f"⏱️ Total time: {elapsed:.1f} seconds ({elapsed/60:.1f} minutes)")
        print(f"📊 {len(contacts)} contacts successfully exported and uploaded")
    else:
        print("\n❌ Upload failed but contacts are saved locally")

if __name__ == "__main__":
    main()