#!/usr/bin/env python3
"""
Simple Export - Basic fields only for speed
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

def export_all_simple():
    """Export all contacts with basic fields only - SINGLE SCRIPT"""
    print("📦 Exporting all 3,792 contacts (basic fields)...")
    print("⏳ This should take about 30-60 seconds...")

    script = '''
    tell application "Contacts"
        set output to ""
        set peopleList to people

        repeat with p in peopleList
            set row to ""

            -- First Name
            try
                set fn to first name of p
                set row to row & fn
            on error
                -- empty
            end try
            set row to row & "|"

            -- Last Name
            try
                set ln to last name of p
                set row to row & ln
            on error
                -- empty
            end try
            set row to row & "|"

            -- Organization
            try
                set org to organization of p
                set row to row & org
            on error
                -- empty
            end try
            set row to row & "|"

            -- First email only
            try
                set emailList to emails of p
                if (count of emailList) > 0 then
                    set row to row & (value of item 1 of emailList)
                end if
            on error
                -- empty
            end try
            set row to row & "|"

            -- First phone only
            try
                set phoneList to phones of p
                if (count of phoneList) > 0 then
                    set row to row & (value of item 1 of phoneList)
                end if
            on error
                -- empty
            end try

            -- Only add if has some data
            if row is not "||||" then
                if output is not "" then set output to output & "\\n"
                set output to output & row
            end if
        end repeat

        return output
    end tell
    '''

    try:
        print("   Running AppleScript...")
        result = subprocess.run(['osascript', '-e', script],
                              capture_output=True, text=True, timeout=120)

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
                            'Primary Email': parts[3],
                            'Primary Phone': parts[4] if len(parts) > 4 else ''
                        }
                        # Add if has any data
                        if any(contact.values()):
                            contacts.append(contact)

            print(f"✅ Extracted {len(contacts)} contacts")
            return contacts
        else:
            print(f"❌ AppleScript error")
            return []

    except subprocess.TimeoutExpired:
        print("❌ Timed out after 2 minutes")
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

        # Upload all data
        print(f"   Uploading {len(contacts)} rows...")
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body={'values': values}
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
    print("SIMPLE EXPORT - ALL 3,792 CONTACTS")
    print("Basic fields only for speed")
    print("="*60)

    start_time = time.time()

    # Export all contacts
    contacts = export_all_simple()

    if not contacts:
        print("❌ No contacts exported")
        sys.exit(1)

    # Save backup
    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"simple_all_{timestamp}.json"

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts, f, indent=2, ensure_ascii=False)
    print(f"💾 Backup saved: {json_file}")

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