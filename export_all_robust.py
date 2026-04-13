#!/usr/bin/env python3
"""
Robust export for ALL contacts with detailed fields and numbered columns
Uses simplified AppleScript calls to avoid timeouts
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

def get_contact_comprehensive(idx):
    """Get single contact with ALL fields - simplified approach"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {idx}
            set output to ""

            -- Basic Info
            try
                set fn to first name of p
                if fn is missing value then set fn to ""
                set output to output & fn
            on error
                set output to output & ""
            end try
            set output to output & "|"

            try
                set ln to last name of p
                if ln is missing value then set ln to ""
                set output to output & ln
            on error
                set output to output & ""
            end try
            set output to output & "|"

            try
                set mn to middle name of p
                if mn is missing value then set mn to ""
                set output to output & mn
            on error
                set output to output & ""
            end try
            set output to output & "|"

            try
                set nn to nickname of p
                if nn is missing value then set nn to ""
                set output to output & nn
            on error
                set output to output & ""
            end try
            set output to output & "|"

            try
                set org to organization of p
                if org is missing value then set org to ""
                set output to output & org
            on error
                set output to output & ""
            end try
            set output to output & "|"

            try
                set job to job title of p
                if job is missing value then set job to ""
                set output to output & job
            on error
                set output to output & ""
            end try
            set output to output & "|"

            -- Birthday
            try
                set bd to birth date of p
                if bd is not missing value then
                    set output to output & (month of bd as string) & "/" & (day of bd as string) & "/" & (year of bd as string)
                end if
            on error
                -- empty
            end try
            set output to output & "|"

            -- ALL Emails
            set emailData to ""
            try
                set emailList to emails of p
                repeat with i from 1 to (count of emailList)
                    set emailVal to value of item i of emailList
                    set emailLabel to "other"
                    try
                        set labelVal to label of item i of emailList
                        if labelVal is not missing value then
                            set labelStr to labelVal as string
                            if labelStr contains "Work" then
                                set emailLabel to "work"
                            else if labelStr contains "Home" then
                                set emailLabel to "home"
                            end if
                        end if
                    end try

                    if emailData is not "" then set emailData to emailData & ";"
                    set emailData to emailData & emailLabel & ":" & emailVal
                end repeat
            end try
            set output to output & emailData & "|"

            -- ALL Phones
            set phoneData to ""
            try
                set phoneList to phones of p
                repeat with i from 1 to (count of phoneList)
                    set phoneVal to value of item i of phoneList
                    set phoneLabel to "other"
                    try
                        set labelVal to label of item i of phoneList
                        if labelVal is not missing value then
                            set labelStr to labelVal as string
                            if labelStr contains "Mobile" or labelStr contains "iPhone" then
                                set phoneLabel to "mobile"
                            else if labelStr contains "Work" and labelStr contains "FAX" then
                                set phoneLabel to "workfax"
                            else if labelStr contains "Home" and labelStr contains "FAX" then
                                set phoneLabel to "homefax"
                            else if labelStr contains "Work" then
                                set phoneLabel to "work"
                            else if labelStr contains "Home" then
                                set phoneLabel to "home"
                            end if
                        end if
                    end try

                    if phoneData is not "" then set phoneData to phoneData & ";"
                    set phoneData to phoneData & phoneLabel & ":" & phoneVal
                end repeat
            end try
            set output to output & phoneData & "|"

            -- ALL Addresses
            set addrData to ""
            try
                set addrList to addresses of p
                repeat with i from 1 to (count of addrList)
                    set addr to item i of addrList
                    set addrLabel to "other"
                    try
                        set labelVal to label of item i of addrList
                        if labelVal is not missing value then
                            set labelStr to labelVal as string
                            if labelStr contains "Work" then
                                set addrLabel to "work"
                            else if labelStr contains "Home" then
                                set addrLabel to "home"
                            end if
                        end if
                    end try

                    set addrParts to ""
                    try
                        set streetVal to street of addr
                        if streetVal is not missing value then set addrParts to streetVal as string
                    end try
                    try
                        set cityVal to city of addr
                        if cityVal is not missing value then
                            if addrParts is not "" then set addrParts to addrParts & ", "
                            set addrParts to addrParts & (cityVal as string)
                        end if
                    end try
                    try
                        set stateVal to state of addr
                        if stateVal is not missing value then
                            if addrParts is not "" then set addrParts to addrParts & ", "
                            set addrParts to addrParts & (stateVal as string)
                        end if
                    end try
                    try
                        set zipVal to zip of addr
                        if zipVal is not missing value then
                            if addrParts is not "" then set addrParts to addrParts & " "
                            set addrParts to addrParts & (zipVal as string)
                        end if
                    end try

                    if addrParts is not "" then
                        if addrData is not "" then set addrData to addrData & ";"
                        set addrData to addrData & addrLabel & ":" & addrParts
                    end if
                end repeat
            end try
            set output to output & addrData & "|"

            -- Notes
            try
                set nt to note of p
                if nt is not missing value then
                    set output to output & nt
                end if
            on error
                -- empty
            end try

            return output
        on error
            return "ERROR"
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=8)
        if result.returncode == 0 and not result.stdout.strip().startswith("ERROR"):
            return result.stdout.strip()
    except:
        pass
    return None

def parse_contact_data(contact_data):
    """Parse the contact data into numbered columns"""
    if not contact_data:
        return None

    parts = contact_data.split('|')
    if len(parts) < 11:
        return None

    contact = {
        'First Name': parts[0] if parts[0] != "missing value" else '',
        'Last Name': parts[1] if parts[1] != "missing value" else '',
        'Middle Name': parts[2] if parts[2] != "missing value" else '',
        'Nickname': parts[3] if parts[3] != "missing value" else '',
        'Organization': parts[4] if parts[4] != "missing value" else '',
        'Job Title': parts[5] if parts[5] != "missing value" else '',
        'Birthday': parts[6] if parts[6] != "missing value" else '',
        'Notes': parts[10] if len(parts) > 10 and parts[10] != "missing value" else ''
    }

    # Parse emails into numbered columns
    email_data = parts[7] if len(parts) > 7 else ''
    if email_data:
        email_counts = {'home': 0, 'work': 0, 'other': 0}
        for email_item in email_data.split(';'):
            if ':' in email_item:
                label, value = email_item.split(':', 1)
                email_counts[label] += 1
                contact[f'{label.title()} Email {email_counts[label]}'] = value

    # Parse phones into numbered columns
    phone_data = parts[8] if len(parts) > 8 else ''
    if phone_data:
        phone_counts = {'mobile': 0, 'home': 0, 'work': 0, 'workfax': 0, 'homefax': 0, 'other': 0}
        for phone_item in phone_data.split(';'):
            if ':' in phone_item:
                label, value = phone_item.split(':', 1)
                phone_counts[label] += 1
                if label == 'workfax':
                    contact[f'Work Fax {phone_counts[label]}'] = value
                elif label == 'homefax':
                    contact[f'Home Fax {phone_counts[label]}'] = value
                else:
                    contact[f'{label.title()} Phone {phone_counts[label]}'] = value

    # Parse addresses into numbered columns
    addr_data = parts[9] if len(parts) > 9 else ''
    if addr_data:
        addr_counts = {'home': 0, 'work': 0, 'other': 0}
        for addr_item in addr_data.split(';'):
            if ':' in addr_item:
                label, value = addr_item.split(':', 1)
                addr_counts[label] += 1
                contact[f'{label.title()} Address {addr_counts[label]}'] = value

    return contact

def export_all_robust():
    """Export ALL contacts robustly"""
    # Get total count
    count_script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', count_script], capture_output=True, text=True)
    total = int(result.stdout.strip())

    print(f"📊 Found {total:,} contacts in Mac Contacts app")
    print("🚀 Starting robust export with ALL fields and numbered columns...\n")

    # Check for existing progress
    progress_file = Path('exports/robust_progress.json')
    if progress_file.exists():
        with open(progress_file, 'r') as f:
            data = json.load(f)
            all_contacts = data.get('contacts', [])
            all_columns = set(data.get('columns', []))
            last_idx = data.get('last_index', 0)
            print(f"📂 Resuming from contact {last_idx + 1}...")
            print(f"   Already exported: {len(all_contacts)}\n")
    else:
        all_contacts = []
        all_columns = set()
        last_idx = 0
        progress_file.parent.mkdir(exist_ok=True)

    # Process contacts with batch restarts
    batch_size = 100  # Restart Contacts app every 100 contacts

    for batch_start in range(last_idx + 1, total + 1, batch_size):
        batch_end = min(batch_start + batch_size - 1, total)
        print(f"📦 Processing batch {batch_start}-{batch_end}...")

        # Restart Contacts app for each batch to prevent hangs
        print("   🔄 Restarting Contacts app...")
        subprocess.run(['osascript', '-e', 'tell application "Contacts" to quit'], capture_output=True)
        time.sleep(2)
        subprocess.run(['open', '-a', 'Contacts'], capture_output=True)
        time.sleep(3)

        batch_contacts = 0
        for i in range(batch_start, batch_end + 1):
            try:
                contact_data = get_contact_comprehensive(i)
                if contact_data:
                    contact = parse_contact_data(contact_data)
                    if contact:
                        # Check if contact has any meaningful data
                        has_data = any(v and v.strip() and v != "missing value" for v in contact.values())
                        if has_data:
                            all_columns.update(contact.keys())
                            all_contacts.append(contact)
                            batch_contacts += 1
            except Exception as e:
                print(f"   ⚠️ Error with contact {i}: {e}")
                continue

        print(f"   ✅ Batch complete: +{batch_contacts} contacts (Total: {len(all_contacts)})")

        # Save progress after each batch
        with open(progress_file, 'w') as f:
            json.dump({
                'contacts': all_contacts,
                'columns': list(all_columns),
                'last_index': batch_end
            }, f)
        print(f"   💾 Progress saved: {len(all_contacts)} contacts, {len(all_columns)} columns")

    # Final save
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_file = progress_file.parent / f"robust_export_{timestamp}.json"
    with open(final_file, 'w') as f:
        json.dump(all_contacts, f, indent=2)

    print(f"\n🎉 ROBUST EXPORT COMPLETE!")
    print(f"✅ Total contacts exported: {len(all_contacts)}")
    print(f"📊 Total unique columns: {len(all_columns)}")
    print(f"💾 Final backup saved: {final_file}")

    # Clean up progress file
    if progress_file.exists():
        progress_file.unlink()

    return all_contacts, sorted(all_columns)

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

def upload_robust_to_sheets(service, contacts_data):
    """Upload robust export to Google Sheets"""
    contacts, all_columns = contacts_data
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')

    if not contacts:
        print("❌ No contacts to upload")
        return False

    print(f"\n📊 Uploading {len(contacts)} contacts with {len(all_columns)} columns...")

    # Build dynamic headers
    base_headers = [
        'First Name', 'Last Name', 'Middle Name', 'Nickname',
        'Organization', 'Job Title', 'Birthday'
    ]

    # Collect dynamic headers
    email_headers = sorted([col for col in all_columns if 'Email' in col])
    phone_headers = sorted([col for col in all_columns if 'Phone' in col or 'Fax' in col])
    address_headers = sorted([col for col in all_columns if 'Address' in col])
    other_headers = ['Notes']

    headers = base_headers + email_headers + phone_headers + address_headers + other_headers

    # Build values array
    values = [headers]
    for contact in contacts:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear sheet
        print("   Clearing sheet...")
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:ZZ"
        ).execute()

        # Upload in chunks
        chunk_size = 2000
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

        # Format sheet
        print("   Formatting...")
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

        print(f"✅ Successfully uploaded {len(contacts)} contacts with {len(headers)} columns!")
        print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return True

    except Exception as e:
        print(f"❌ Upload error: {e}")
        return False

def main():
    print("="*60)
    print("ROBUST EXPORT - ALL CONTACTS WITH NUMBERED COLUMNS")
    print("Restarts Contacts app every 100 contacts to prevent hangs")
    print("="*60)

    start_time = time.time()

    # Export all contacts robustly
    contacts_data = export_all_robust()

    if not contacts_data[0]:
        print("❌ No contacts exported")
        sys.exit(1)

    # Upload to Google Sheets
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    if upload_robust_to_sheets(service, contacts_data):
        elapsed = time.time() - start_time
        print(f"\n🎉 COMPLETE!")
        print(f"⏱️ Total time: {elapsed:.1f} seconds ({elapsed/60:.1f} minutes)")
        print(f"📊 {len(contacts_data[0])} contacts with {len(contacts_data[1])} fields")
        print("🔥 ALL contacts exported with numbered columns!")
    else:
        print("\n❌ Upload failed but contacts are saved locally")

if __name__ == "__main__":
    main()