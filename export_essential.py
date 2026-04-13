#!/usr/bin/env python3
"""
Essential Export - Get basic fields for all 3,792 contacts FAST
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

def export_all_essential():
    """Export all contacts - names and primary contact info only"""
    print("📦 Exporting all 3,792 contacts (essential fields only)...")
    print("⏳ Processing in small batches for speed...\n")

    all_contacts = []
    batch_size = 50  # Small batches for speed

    # Get total count first
    count_script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', count_script], capture_output=True, text=True)
    total = int(result.stdout.strip())
    print(f"📊 Total contacts: {total:,}\n")

    for start in range(1, total + 1, batch_size):
        end = min(start + batch_size - 1, total)

        batch_script = f'''
        tell application "Contacts"
            set output to ""
            repeat with i from {start} to {end}
                try
                    set p to person i
                    set row to ""

                    -- Get first name
                    try
                        set row to row & (first name of p)
                    end try
                    set row to row & "|"

                    -- Get last name
                    try
                        set row to row & (last name of p)
                    end try
                    set row to row & "|"

                    -- Get organization
                    try
                        set row to row & (organization of p)
                    end try
                    set row to row & "|"

                    -- Get first email
                    try
                        set emailList to emails of p
                        if (count of emailList) > 0 then
                            set row to row & (value of item 1 of emailList)
                        end if
                    end try
                    set row to row & "|"

                    -- Get first phone
                    try
                        set phoneList to phones of p
                        if (count of phoneList) > 0 then
                            set row to row & (value of item 1 of phoneList)
                        end if
                    end try

                    set output to output & row & "\\n"
                on error
                    -- Skip problematic contact
                end try
            end repeat
            return output
        end tell
        '''

        try:
            result = subprocess.run(['osascript', '-e', batch_script],
                                  capture_output=True, text=True, timeout=10)

            if result.returncode == 0:
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
                            # Add if has any data
                            if any(contact.values()):
                                all_contacts.append(contact)

            # Progress update
            percent = (end / total) * 100
            print(f"✅ Batch {start}-{end}: {len(all_contacts)} contacts ({percent:.1f}%)")

        except subprocess.TimeoutExpired:
            print(f"⚠️ Batch {start}-{end} timed out, skipping...")
            continue
        except Exception as e:
            print(f"⚠️ Batch {start}-{end} error: {e}")
            continue

    print(f"\n✅ Total exported: {len(all_contacts)} contacts")
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

        # Upload in chunks
        chunk_size = 5000
        for i in range(0, len(values), chunk_size):
            chunk = values[i:i + chunk_size]
            range_start = f'A{i+1}' if i == 0 else f'A{i+1}'

            print(f"   Uploading rows {i+1} to {min(i+chunk_size, len(values))}...")
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
    print("ESSENTIAL EXPORT - ALL 3,792 CONTACTS")
    print("Basic fields for maximum speed")
    print("="*60)

    start_time = time.time()

    # Export all contacts
    contacts = export_all_essential()

    if not contacts:
        print("❌ No contacts exported")
        sys.exit(1)

    # Save backup
    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"essential_all_{timestamp}.json"

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts, f, indent=2, ensure_ascii=False)
    print(f"\n💾 Backup saved: {json_file}")

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