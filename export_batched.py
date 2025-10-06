#!/usr/bin/env python3
"""
Batch Export - Process contacts in chunks
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

def get_total_contacts():
    """Get total number of contacts"""
    script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    return int(result.stdout.strip())

def export_batch(start_idx, end_idx):
    """Export a batch of contacts"""
    script = f'''
    tell application "Contacts"
        set output to ""
        repeat with i from {start_idx} to {end_idx}
            try
                set p to person i
                set row to ""

                -- First Name
                try
                    set row to row & (first name of p)
                on error
                    -- empty
                end try
                set row to row & "|"

                -- Last Name
                try
                    set row to row & (last name of p)
                on error
                    -- empty
                end try
                set row to row & "|"

                -- Organization
                try
                    set row to row & (organization of p)
                on error
                    -- empty
                end try
                set row to row & "|"

                -- First email
                try
                    set emailList to emails of p
                    if (count of emailList) > 0 then
                        set row to row & (value of item 1 of emailList)
                    end if
                on error
                    -- empty
                end try
                set row to row & "|"

                -- First phone
                try
                    set phoneList to phones of p
                    if (count of phoneList) > 0 then
                        set row to row & (value of item 1 of phoneList)
                    end if
                on error
                    -- empty
                end try

                if row is not "||||" then
                    set output to output & row & "\\n"
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
                              capture_output=True, text=True, timeout=30)
        if result.returncode == 0:
            contacts = []
            lines = result.stdout.strip().split('\n')
            for line in lines:
                if line:
                    parts = line.split('|')
                    if len(parts) >= 5:
                        contact = {
                            'First Name': parts[0],
                            'Last Name': parts[1],
                            'Organization': parts[2],
                            'Email': parts[3],
                            'Phone': parts[4] if len(parts) > 4 else ''
                        }
                        if any(contact.values()):
                            contacts.append(contact)
            return contacts
    except:
        return []

def export_all_contacts():
    """Export all contacts in batches"""
    total = get_total_contacts()
    print(f"📊 Found {total:,} contacts")

    all_contacts = []
    batch_size = 100

    for start in range(1, total + 1, batch_size):
        end = min(start + batch_size - 1, total)
        print(f"📦 Exporting batch {start}-{end}...")

        batch_contacts = export_batch(start, end)
        all_contacts.extend(batch_contacts)

        # Show progress
        percent = (end / total) * 100
        print(f"   ✅ {len(all_contacts)} contacts exported ({percent:.1f}%)")

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

    headers = ['First Name', 'Last Name', 'Organization', 'Email', 'Phone']
    values = [headers]

    for contact in contacts:
        row = [contact.get(h, '') for h in headers]
        values.append(row)

    try:
        # Clear sheet
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Upload all data
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        print(f"✅ Uploaded {len(contacts)} contacts!")
        print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return True
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    print("="*60)
    print("BATCH EXPORT - ALL CONTACTS")
    print("="*60)

    start_time = time.time()

    # Export all contacts
    contacts = export_all_contacts()

    if not contacts:
        print("❌ No contacts exported")
        sys.exit(1)

    # Save backup
    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"batch_export_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts, f, indent=2, ensure_ascii=False)
    print(f"💾 Backup saved: {json_file}")

    # Upload to Google Sheets
    print("\n🔐 Authenticating...")
    service = authenticate_google_sheets()

    if upload_to_sheets(service, contacts):
        elapsed = time.time() - start_time
        print(f"\n🎉 COMPLETE!")
        print(f"⏱️ Time: {elapsed:.1f} seconds")
        print(f"📊 {len(contacts)} contacts exported")

if __name__ == "__main__":
    main()