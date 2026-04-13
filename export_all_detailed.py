#!/usr/bin/env python3
"""
Export ALL contacts with ALL FIELDS and numbered columns
Based on the working test_3_all_fields.py script
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

def get_field_safely(person_var, field_name, script_prefix=""):
    """Get a field value safely from AppleScript"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {person_var}
            {script_prefix}
            set val to {field_name} of p
            if val is missing value then
                return ""
            else
                return val as string
            end if
        on error
            return ""
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=2)
        if result.returncode == 0:
            return result.stdout.strip()
    except:
        pass
    return ""

def get_emails_separated(person_idx):
    """Get emails as individual items"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {person_idx}
            set homeEmails to {{}}
            set workEmails to {{}}
            set otherEmails to {{}}
            set emailList to emails of p
            if (count of emailList) > 0 then
                repeat with i from 1 to (count of emailList)
                    set emailVal to value of item i of emailList
                    set emailLabel to "other"
                    try
                        set labelVal to label of item i of emailList
                        if labelVal is not missing value then
                            set labelStr to labelVal as string
                            -- Clean up internal labels
                            if labelStr contains "Work" then
                                set emailLabel to "work"
                            else if labelStr contains "Home" then
                                set emailLabel to "home"
                            else if labelStr contains "Other" or labelStr contains "_$!<" then
                                set emailLabel to "other"
                            end if
                        end if
                    end try

                    -- Split multiple emails that might be semicolon-separated
                    set emailValues to {{}}
                    if emailVal contains ";" then
                        set AppleScript's text item delimiters to ";"
                        set emailValues to text items of emailVal
                        set AppleScript's text item delimiters to ""
                    else
                        set emailValues to {{emailVal}}
                    end if

                    repeat with singleEmail in emailValues
                        set trimmedEmail to singleEmail
                        -- Trim leading/trailing spaces
                        repeat while trimmedEmail starts with " "
                            set trimmedEmail to text 2 thru -1 of trimmedEmail
                        end repeat
                        repeat while trimmedEmail ends with " "
                            set trimmedEmail to text 1 thru -2 of trimmedEmail
                        end repeat

                        if emailLabel is "home" then
                            set end of homeEmails to trimmedEmail
                        else if emailLabel is "work" then
                            set end of workEmails to trimmedEmail
                        else
                            set end of otherEmails to trimmedEmail
                        end if
                    end repeat
                end repeat
            end if

            set output to ""
            -- Output home emails
            repeat with i from 1 to (count of homeEmails)
                if output is not "" then set output to output & "|"
                set output to output & "home:" & item i of homeEmails
            end repeat
            -- Output work emails
            repeat with i from 1 to (count of workEmails)
                if output is not "" then set output to output & "|"
                set output to output & "work:" & item i of workEmails
            end repeat
            -- Output other emails
            repeat with i from 1 to (count of otherEmails)
                if output is not "" then set output to output & "|"
                set output to output & "other:" & item i of otherEmails
            end repeat

            return output
        on error
            return ""
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            emails = {'home': [], 'work': [], 'other': []}
            if result.stdout.strip():
                parts = result.stdout.strip().split('|')
                for part in parts:
                    if ':' in part:
                        type_label, value = part.split(':', 1)
                        if type_label in emails:
                            emails[type_label].append(value)
            return emails
    except:
        pass
    return {'home': [], 'work': [], 'other': []}

def get_phones_separated(person_idx):
    """Get phones as individual items"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {person_idx}
            set mobilePhones to {{}}
            set homePhones to {{}}
            set workPhones to {{}}
            set workFaxes to {{}}
            set homeFaxes to {{}}
            set otherPhones to {{}}
            set phoneList to phones of p
            if (count of phoneList) > 0 then
                repeat with i from 1 to (count of phoneList)
                    set phoneVal to value of item i of phoneList
                    set phoneType to "other"
                    try
                        set labelVal to label of item i of phoneList
                        if labelVal is not missing value then
                            set labelStr to labelVal as string
                            -- Clean up internal labels and categorize
                            if labelStr contains "Mobile" or labelStr contains "iPhone" then
                                set phoneType to "mobile"
                            else if labelStr contains "Work" and labelStr contains "FAX" then
                                set phoneType to "workfax"
                            else if labelStr contains "Home" and labelStr contains "FAX" then
                                set phoneType to "homefax"
                            else if labelStr contains "Work" then
                                set phoneType to "work"
                            else if labelStr contains "Home" then
                                set phoneType to "home"
                            else if labelStr contains "Main" then
                                set phoneType to "work"
                            else if labelStr contains "Other" or labelStr contains "_$!<" then
                                set phoneType to "other"
                            end if
                        end if
                    end try

                    -- Split multiple numbers that are semicolon-separated
                    set phoneNumbers to {{}}
                    if phoneVal contains ";" then
                        set AppleScript's text item delimiters to ";"
                        set phoneNumbers to text items of phoneVal
                        set AppleScript's text item delimiters to ""
                    else
                        set phoneNumbers to {{phoneVal}}
                    end if

                    repeat with phoneNum in phoneNumbers
                        set trimmedPhone to phoneNum
                        -- Trim leading/trailing spaces
                        repeat while trimmedPhone starts with " "
                            set trimmedPhone to text 2 thru -1 of trimmedPhone
                        end repeat
                        repeat while trimmedPhone ends with " "
                            set trimmedPhone to text 1 thru -2 of trimmedPhone
                        end repeat

                        if phoneType is "mobile" then
                            set end of mobilePhones to trimmedPhone
                        else if phoneType is "home" then
                            set end of homePhones to trimmedPhone
                        else if phoneType is "work" then
                            set end of workPhones to trimmedPhone
                        else if phoneType is "workfax" then
                            set end of workFaxes to trimmedPhone
                        else if phoneType is "homefax" then
                            set end of homeFaxes to trimmedPhone
                        else
                            set end of otherPhones to trimmedPhone
                        end if
                    end repeat
                end repeat
            end if

            set output to ""
            -- Output mobile phones
            repeat with i from 1 to (count of mobilePhones)
                if output is not "" then set output to output & "|"
                set output to output & "mobile:" & item i of mobilePhones
            end repeat
            -- Output home phones
            repeat with i from 1 to (count of homePhones)
                if output is not "" then set output to output & "|"
                set output to output & "home:" & item i of homePhones
            end repeat
            -- Output work phones
            repeat with i from 1 to (count of workPhones)
                if output is not "" then set output to output & "|"
                set output to output & "work:" & item i of workPhones
            end repeat
            -- Output work faxes
            repeat with i from 1 to (count of workFaxes)
                if output is not "" then set output to output & "|"
                set output to output & "workfax:" & item i of workFaxes
            end repeat
            -- Output home faxes
            repeat with i from 1 to (count of homeFaxes)
                if output is not "" then set output to output & "|"
                set output to output & "homefax:" & item i of homeFaxes
            end repeat
            -- Output other phones
            repeat with i from 1 to (count of otherPhones)
                if output is not "" then set output to output & "|"
                set output to output & "other:" & item i of otherPhones
            end repeat

            return output
        on error
            return ""
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            phones = {'mobile': [], 'home': [], 'work': [], 'workfax': [], 'homefax': [], 'other': []}
            if result.stdout.strip():
                parts = result.stdout.strip().split('|')
                for part in parts:
                    if ':' in part:
                        type_label, value = part.split(':', 1)
                        if type_label in phones:
                            phones[type_label].append(value)
            return phones
    except:
        pass
    return {'mobile': [], 'home': [], 'work': [], 'workfax': [], 'homefax': [], 'other': []}

def get_addresses_separated(person_idx):
    """Get addresses as individual items"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {person_idx}
            set homeAddresses to {{}}
            set workAddresses to {{}}
            set otherAddresses to {{}}
            set addrList to addresses of p
            if (count of addrList) > 0 then
                repeat with i from 1 to (count of addrList)
                    set addr to item i of addrList
                    set addrType to "other"
                    try
                        set labelVal to label of item i of addrList
                        if labelVal is not missing value then
                            set labelStr to labelVal as string
                            -- Clean up internal labels
                            if labelStr contains "Work" then
                                set addrType to "work"
                            else if labelStr contains "Home" then
                                set addrType to "home"
                            else if labelStr contains "Other" or labelStr contains "_$!<" then
                                set addrType to "other"
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
                    try
                        set countryVal to country of addr
                        if countryVal is not missing value then
                            if addrParts is not "" then set addrParts to addrParts & ", "
                            set addrParts to addrParts & (countryVal as string)
                        end if
                    end try

                    if addrParts is not "" then
                        if addrType is "home" then
                            set end of homeAddresses to addrParts
                        else if addrType is "work" then
                            set end of workAddresses to addrParts
                        else
                            set end of otherAddresses to addrParts
                        end if
                    end if
                end repeat
            end if

            set output to ""
            -- Output home addresses
            repeat with i from 1 to (count of homeAddresses)
                if output is not "" then set output to output & "|"
                set output to output & "home:" & item i of homeAddresses
            end repeat
            -- Output work addresses
            repeat with i from 1 to (count of workAddresses)
                if output is not "" then set output to output & "|"
                set output to output & "work:" & item i of workAddresses
            end repeat
            -- Output other addresses
            repeat with i from 1 to (count of otherAddresses)
                if output is not "" then set output to output & "|"
                set output to output & "other:" & item i of otherAddresses
            end repeat

            return output
        on error
            return ""
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            addresses = {'home': [], 'work': [], 'other': []}
            if result.stdout.strip():
                parts = result.stdout.strip().split('|')
                for part in parts:
                    if ':' in part:
                        type_label, value = part.split(':', 1)
                        if type_label in addresses:
                            addresses[type_label].append(value)
            return addresses
    except:
        pass
    return {'home': [], 'work': [], 'other': []}

def export_contact_all_fields(idx):
    """Export single contact with all available fields in separate columns"""
    contact = {}

    # Names
    contact['First Name'] = get_field_safely(idx, 'first name')
    contact['Last Name'] = get_field_safely(idx, 'last name')
    contact['Middle Name'] = get_field_safely(idx, 'middle name')
    contact['Nickname'] = get_field_safely(idx, 'nickname')
    contact['Name Prefix'] = get_field_safely(idx, 'name prefix')
    contact['Name Suffix'] = get_field_safely(idx, 'name suffix')

    # Organization
    contact['Organization'] = get_field_safely(idx, 'organization')
    contact['Job Title'] = get_field_safely(idx, 'job title')
    contact['Department'] = get_field_safely(idx, 'department')

    # Birthday
    birthday_script = f'''
    tell application "Contacts"
        try
            set p to person {idx}
            set bd to birth date of p
            if bd is not missing value then
                return (month of bd as string) & "/" & (day of bd as string) & "/" & (year of bd as string)
            else
                return ""
            end if
        on error
            return ""
        end try
    end tell
    '''
    try:
        result = subprocess.run(['osascript', '-e', birthday_script], capture_output=True, text=True, timeout=2)
        contact['Birthday'] = result.stdout.strip() if result.returncode == 0 else ""
    except:
        contact['Birthday'] = ""

    # Emails separated into individual columns
    emails = get_emails_separated(idx)
    # Store each email in numbered columns
    for i, email in enumerate(emails.get('home', []), 1):
        contact[f'Home Email {i}'] = email
    for i, email in enumerate(emails.get('work', []), 1):
        contact[f'Work Email {i}'] = email
    for i, email in enumerate(emails.get('other', []), 1):
        contact[f'Other Email {i}'] = email

    # Phones separated into individual columns
    phones = get_phones_separated(idx)
    # Store each phone in numbered columns
    for i, phone in enumerate(phones.get('mobile', []), 1):
        contact[f'Mobile Phone {i}'] = phone
    for i, phone in enumerate(phones.get('home', []), 1):
        contact[f'Home Phone {i}'] = phone
    for i, phone in enumerate(phones.get('work', []), 1):
        contact[f'Work Phone {i}'] = phone
    for i, fax in enumerate(phones.get('workfax', []), 1):
        contact[f'Work Fax {i}'] = fax
    for i, fax in enumerate(phones.get('homefax', []), 1):
        contact[f'Home Fax {i}'] = fax
    for i, phone in enumerate(phones.get('other', []), 1):
        contact[f'Other Phone {i}'] = phone

    # Addresses separated into individual columns
    addresses = get_addresses_separated(idx)
    # Store each address in numbered columns
    for i, address in enumerate(addresses.get('home', []), 1):
        contact[f'Home Address {i}'] = address
    for i, address in enumerate(addresses.get('work', []), 1):
        contact[f'Work Address {i}'] = address
    for i, address in enumerate(addresses.get('other', []), 1):
        contact[f'Other Address {i}'] = address

    # Notes
    contact['Notes'] = get_field_safely(idx, 'note')

    # URLs - get all and number them
    url_script = f'''
    tell application "Contacts"
        try
            set p to person {idx}
            set output to ""
            set urlList to urls of p
            if (count of urlList) > 0 then
                repeat with i from 1 to (count of urlList)
                    if i > 1 then set output to output & "|"
                    set output to output & value of item i of urlList
                end repeat
            end if
            return output
        on error
            return ""
        end try
    end tell
    '''
    try:
        result = subprocess.run(['osascript', '-e', url_script], capture_output=True, text=True, timeout=2)
        if result.returncode == 0 and result.stdout.strip():
            urls = result.stdout.strip().split('|')
            for i, url in enumerate(urls, 1):
                contact[f'URL {i}'] = url
    except:
        pass

    return contact

def export_all_contacts_detailed():
    """Export ALL contacts with detailed fields, 3 at a time"""
    # Get total count
    count_script = 'tell application "Contacts" to count of people'
    result = subprocess.run(['osascript', '-e', count_script], capture_output=True, text=True)
    total = int(result.stdout.strip())

    print(f"📊 Found {total:,} contacts in Mac Contacts app")
    print("🚀 Starting detailed export with numbered columns...\n")

    # Check for existing progress
    progress_file = Path('exports/detailed_progress.json')
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

    # Process contacts one by one (safer than batches for detailed extraction)
    batch_count = 0
    for i in range(last_idx + 1, total + 1):
        try:
            contact = export_contact_all_fields(i)
            if contact:
                # More lenient filtering - include if any field has meaningful data
                has_data = False
                for key, value in contact.items():
                    if value and value != "missing value" and value.strip():
                        has_data = True
                        break

                if has_data:
                    all_columns.update(contact.keys())
                    all_contacts.append(contact)
        except Exception as e:
            print(f"   ⚠️ Error with contact {i}: {e}")
            continue

        # Progress update every 10 contacts
        if i % 10 == 0:
            percent = (i / total) * 100
            print(f"✅ Processed {i}/{total} ({percent:.1f}%) - Found {len(all_contacts)} valid contacts")

        # Save progress every 50 contacts
        if i % 50 == 0:
            with open(progress_file, 'w') as f:
                json.dump({
                    'contacts': all_contacts,
                    'columns': list(all_columns),
                    'last_index': i
                }, f)
            print(f"   💾 Progress saved: {len(all_contacts)} contacts, {len(all_columns)} columns")

        # Brief pause every 100 contacts
        if i % 100 == 0:
            print("   ⏸️ Brief pause...")
            time.sleep(1)

    # Final save
    with open(progress_file, 'w') as f:
        json.dump({
            'contacts': all_contacts,
            'columns': list(all_columns),
            'last_index': total
        }, f)

    # Save final backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_file = progress_file.parent / f"detailed_export_{timestamp}.json"
    with open(final_file, 'w') as f:
        json.dump(all_contacts, f, indent=2)

    print(f"\n🎉 DETAILED EXPORT COMPLETE!")
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

def upload_detailed_to_sheets(service, contacts_data):
    """Upload detailed contacts to Google Sheets with numbered columns"""
    contacts, all_columns = contacts_data
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')

    if not contacts:
        print("❌ No contacts to upload")
        return False

    print(f"\n📊 Uploading {len(contacts)} contacts with {len(all_columns)} columns...")

    # Build dynamic headers based on all columns found
    base_headers = [
        'First Name', 'Last Name', 'Middle Name', 'Nickname',
        'Name Prefix', 'Name Suffix',
        'Organization', 'Job Title', 'Department',
        'Birthday'
    ]

    # Collect all dynamic field names
    email_headers = sorted([col for col in all_columns if 'Email' in col])
    phone_headers = sorted([col for col in all_columns if 'Phone' in col or 'Fax' in col])
    address_headers = sorted([col for col in all_columns if 'Address' in col])
    url_headers = sorted([col for col in all_columns if col.startswith('URL')])
    other_headers = ['Notes'] + (url_headers if url_headers else [])

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

        # Upload in chunks for large datasets
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
    print("DETAILED EXPORT - ALL CONTACTS WITH ALL FIELDS")
    print("Numbered columns for multiple values")
    print("="*60)

    start_time = time.time()

    # Export all contacts with detailed fields
    contacts_data = export_all_contacts_detailed()

    if not contacts_data[0]:
        print("❌ No contacts exported")
        sys.exit(1)

    # Upload to Google Sheets
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    if upload_detailed_to_sheets(service, contacts_data):
        elapsed = time.time() - start_time
        print(f"\n🎉 COMPLETE!")
        print(f"⏱️ Total time: {elapsed:.1f} seconds ({elapsed/60:.1f} minutes)")
        print(f"📊 {len(contacts_data[0])} contacts with {len(contacts_data[1])} fields")
    else:
        print("\n❌ Upload failed but contacts are saved locally")

if __name__ == "__main__":
    main()