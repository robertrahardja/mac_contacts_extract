#!/usr/bin/env python3
"""
Incremental Export - Process contacts 10 at a time and save progress
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

def get_contact_simple(idx):
    """Get single contact quickly"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {idx}
            set output to ""

            try
                set output to output & (first name of p)
            end try
            set output to output & "|"

            try
                set output to output & (last name of p)
            end try
            set output to output & "|"

            try
                set output to output & (organization of p)
            end try
            set output to output & "|"

            try
                set emailList to emails of p
                if (count of emailList) > 0 then
                    set output to output & (value of item 1 of emailList)
                end if
            end try
            set output to output & "|"

            try
                set phoneList to phones of p
                if (count of phoneList) > 0 then
                    set output to output & (value of item 1 of phoneList)
                end if
            end try

            return output
        on error
            return "||||"
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script],
                              capture_output=True, text=True, timeout=2)
        if result.returncode == 0:
            parts = result.stdout.strip().split('|')
            if len(parts) >= 5 and parts != ['', '', '', '', '']:
                return {
                    'First Name': parts[0],
                    'Last Name': parts[1],
                    'Organization': parts[2],
                    'Primary Email': parts[3],
                    'Primary Phone': parts[4] if len(parts) > 4 else ''
                }
    except:
        pass
    return None

def export_incremental():
    """Export contacts incrementally"""
    print("📦 Starting incremental export...")
    print("⏳ Processing 10 contacts at a time to avoid timeout...\n")

    # Get total count
    count_script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', count_script], capture_output=True, text=True)
    total = int(result.stdout.strip())
    print(f"📊 Total contacts: {total:,}\n")

    # Check for existing progress
    progress_file = Path('exports/incremental_progress.json')
    if progress_file.exists():
        with open(progress_file, 'r') as f:
            data = json.load(f)
            contacts = data['contacts']
            last_idx = data['last_index']
            print(f"📂 Resuming from contact {last_idx + 1}...")
            print(f"   Already exported: {len(contacts)}\n")
    else:
        contacts = []
        last_idx = 0
        progress_file.parent.mkdir(exist_ok=True)

    # Process in tiny batches
    batch_size = 10
    save_every = 50  # Save progress every 50 contacts

    for i in range(last_idx + 1, min(last_idx + 501, total + 1)):  # Process 500 at a time max
        contact = get_contact_simple(i)
        if contact and any(contact.values()):
            contacts.append(contact)

        # Progress indicator
        if i % batch_size == 0:
            percent = (i / total) * 100
            print(f"✅ Processed {i}/{total} ({percent:.1f}%) - Found {len(contacts)} valid contacts")

        # Save progress periodically
        if i % save_every == 0:
            with open(progress_file, 'w') as f:
                json.dump({'contacts': contacts, 'last_index': i}, f)
            print(f"   💾 Progress saved at contact {i}")

        # Stop at 500 for this run to avoid timeout
        if i >= last_idx + 500:
            print(f"\n⏸️ Pausing at contact {i} to avoid timeout")
            print(f"   Run the script again to continue from contact {i + 1}")
            break

    # Final save
    with open(progress_file, 'w') as f:
        json.dump({'contacts': contacts, 'last_index': i}, f)

    print(f"\n✅ This batch complete: {len(contacts)} total contacts exported")

    if i >= total:
        print("🎉 ALL CONTACTS EXPORTED!")
        # Save final backup
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_file = progress_file.parent / f"final_export_{timestamp}.json"
        with open(final_file, 'w') as f:
            json.dump(contacts, f, indent=2)
        print(f"💾 Final backup: {final_file}")

        return contacts, True  # True = complete
    else:
        return contacts, False  # False = more to do

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
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        # Format headers
        requests = [{
            'repeatCell': {
                'range': {'sheetId': 0, 'startRowIndex': 0, 'endRowIndex': 1},
                'cell': {'userEnteredFormat': {'textFormat': {'bold': True}}},
                'fields': 'userEnteredFormat.textFormat.bold'
            }
        }]

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
    print("INCREMENTAL EXPORT - AVOIDING TIMEOUTS")
    print("Processing 500 contacts at a time")
    print("="*60)

    start_time = time.time()

    # Export contacts incrementally
    contacts, is_complete = export_incremental()

    if is_complete:
        # Upload to Google Sheets only if complete
        print("\n🔐 Authenticating with Google...")
        service = authenticate_google_sheets()

        if upload_to_sheets(service, contacts):
            elapsed = time.time() - start_time
            print(f"\n🎉 EXPORT COMPLETE!")
            print(f"⏱️ Total time: {elapsed:.1f} seconds")
            print(f"📊 {len(contacts)} contacts exported")
        else:
            print("\n❌ Upload failed")
    else:
        print("\n⏸️ Partial export complete")
        print("   Run the script again to continue exporting")
        print(f"   Current progress: {len(contacts)} contacts")

if __name__ == "__main__":
    main()