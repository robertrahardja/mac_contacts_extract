#!/usr/bin/env python3
"""
Test 3 contacts with ALL available fields (simplified approach)
"""

import os
import sys
import subprocess
import json
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Google Sheets imports
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

def get_emails_by_type(person_idx):
    """Get emails separated by type"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {person_idx}
            set homeEmail to ""
            set workEmail to ""
            set otherEmail to ""
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

                    if emailLabel is "home" then
                        if homeEmail is not "" then set homeEmail to homeEmail & "; "
                        set homeEmail to homeEmail & emailVal
                    else if emailLabel is "work" then
                        if workEmail is not "" then set workEmail to workEmail & "; "
                        set workEmail to workEmail & emailVal
                    else
                        if otherEmail is not "" then set otherEmail to otherEmail & "; "
                        set otherEmail to otherEmail & emailVal
                    end if
                end repeat
            end if
            return homeEmail & "|" & workEmail & "|" & otherEmail
        on error
            return "|||"
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            parts = result.stdout.strip().split('|')
            return {
                'home': parts[0] if len(parts) > 0 else '',
                'work': parts[1] if len(parts) > 1 else '',
                'other': parts[2] if len(parts) > 2 else ''
            }
    except:
        pass
    return {'home': '', 'work': '', 'other': ''}

def get_phones_by_type(person_idx):
    """Get phones separated by type"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {person_idx}
            set mobilePhone to ""
            set homePhone to ""
            set workPhone to ""
            set workFax to ""
            set homeFax to ""
            set otherPhone to ""
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

                    if phoneType is "mobile" then
                        if mobilePhone is not "" then set mobilePhone to mobilePhone & "; "
                        set mobilePhone to mobilePhone & phoneVal
                    else if phoneType is "home" then
                        if homePhone is not "" then set homePhone to homePhone & "; "
                        set homePhone to homePhone & phoneVal
                    else if phoneType is "work" then
                        if workPhone is not "" then set workPhone to workPhone & "; "
                        set workPhone to workPhone & phoneVal
                    else if phoneType is "workfax" then
                        if workFax is not "" then set workFax to workFax & "; "
                        set workFax to workFax & phoneVal
                    else if phoneType is "homefax" then
                        if homeFax is not "" then set homeFax to homeFax & "; "
                        set homeFax to homeFax & phoneVal
                    else
                        if otherPhone is not "" then set otherPhone to otherPhone & "; "
                        set otherPhone to otherPhone & phoneVal
                    end if
                end repeat
            end if
            return mobilePhone & "|" & homePhone & "|" & workPhone & "|" & workFax & "|" & homeFax & "|" & otherPhone
        on error
            return "|||||"
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            parts = result.stdout.strip().split('|')
            return {
                'mobile': parts[0] if len(parts) > 0 else '',
                'home': parts[1] if len(parts) > 1 else '',
                'work': parts[2] if len(parts) > 2 else '',
                'work_fax': parts[3] if len(parts) > 3 else '',
                'home_fax': parts[4] if len(parts) > 4 else '',
                'other': parts[5] if len(parts) > 5 else ''
            }
    except:
        pass
    return {'mobile': '', 'home': '', 'work': '', 'work_fax': '', 'home_fax': '', 'other': ''}

def get_addresses_by_type(person_idx):
    """Get addresses separated by type"""
    script = f'''
    tell application "Contacts"
        try
            set p to person {person_idx}
            set homeAddr to ""
            set workAddr to ""
            set otherAddr to ""
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

                    if addrType is "home" then
                        if homeAddr is not "" then set homeAddr to homeAddr & "; "
                        set homeAddr to homeAddr & addrParts
                    else if addrType is "work" then
                        if workAddr is not "" then set workAddr to workAddr & "; "
                        set workAddr to workAddr & addrParts
                    else
                        if otherAddr is not "" then set otherAddr to otherAddr & "; "
                        set otherAddr to otherAddr & addrParts
                    end if
                end repeat
            end if
            return homeAddr & "|" & workAddr & "|" & otherAddr
        on error
            return "||"
        end try
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            parts = result.stdout.strip().split('|')
            return {
                'home': parts[0] if len(parts) > 0 else '',
                'work': parts[1] if len(parts) > 1 else '',
                'other': parts[2] if len(parts) > 2 else ''
            }
    except:
        pass
    return {'home': '', 'work': '', 'other': ''}

def export_contact_all_fields(idx):
    """Export single contact with all available fields in separate columns"""
    print(f"   Extracting fields for contact {idx}...")

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

    # Emails by type
    emails = get_emails_by_type(idx)
    contact['Home Email'] = emails['home']
    contact['Work Email'] = emails['work']
    contact['Other Email'] = emails['other']

    # Phones by type
    phones = get_phones_by_type(idx)
    contact['Mobile Phone'] = phones['mobile']
    contact['Home Phone'] = phones['home']
    contact['Work Phone'] = phones['work']
    contact['Work Fax'] = phones['work_fax']
    contact['Home Fax'] = phones['home_fax']
    contact['Other Phone'] = phones['other']

    # Addresses by type
    addresses = get_addresses_by_type(idx)
    contact['Home Address'] = addresses['home']
    contact['Work Address'] = addresses['work']
    contact['Other Address'] = addresses['other']

    # Notes
    contact['Notes'] = get_field_safely(idx, 'note')

    # URLs (simplified - just get first one for now)
    url_script = f'''
    tell application "Contacts"
        try
            set p to person {idx}
            set urlList to urls of p
            if (count of urlList) > 0 then
                return value of item 1 of urlList
            else
                return ""
            end if
        on error
            return ""
        end try
    end tell
    '''
    try:
        result = subprocess.run(['osascript', '-e', url_script], capture_output=True, text=True, timeout=2)
        contact['URLs'] = result.stdout.strip() if result.returncode == 0 else ""
    except:
        contact['URLs'] = ""

    return contact

def export_3_all_fields():
    """Export 3 contacts with all fields"""
    print("🧪 Exporting 3 contacts with ALL available fields...\n")

    contacts = []

    for i in range(1, 4):
        print(f"📱 Processing contact {i}...")
        contact = export_contact_all_fields(i)

        # Show summary
        name = f"{contact['First Name']} {contact['Last Name']}".strip()
        print(f"   ✅ {name or 'Contact ' + str(i)}")
        if contact['Organization']:
            print(f"      Organization: {contact['Organization']}")
        if contact['Work Email'] or contact['Home Email']:
            emails_summary = []
            if contact['Work Email']:
                emails_summary.append(f"Work: {contact['Work Email'][:30]}..." if len(contact['Work Email']) > 30 else f"Work: {contact['Work Email']}")
            if contact['Home Email']:
                emails_summary.append(f"Home: {contact['Home Email'][:30]}..." if len(contact['Home Email']) > 30 else f"Home: {contact['Home Email']}")
            print(f"      Emails: {', '.join(emails_summary)}")
        if contact['Mobile Phone'] or contact['Work Phone']:
            phones_summary = []
            if contact['Mobile Phone']:
                phones_summary.append(f"Mobile: {contact['Mobile Phone']}")
            if contact['Work Phone']:
                phones_summary.append(f"Work: {contact['Work Phone'][:30]}..." if len(contact['Work Phone']) > 30 else f"Work: {contact['Work Phone']}")
            print(f"      Phones: {', '.join(phones_summary)}")

        contacts.append(contact)

    print(f"\n✅ Exported {len(contacts)} contacts with all fields")
    return contacts

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

def upload_all_fields_to_sheets(service, contacts):
    """Upload all fields to Google Sheets with separate columns for each type"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')

    print(f"\n📊 Uploading to Google Sheets with separate columns for each field type...")

    # Define ALL headers with separate columns for each type
    headers = [
        'First Name', 'Last Name', 'Middle Name', 'Nickname',
        'Name Prefix', 'Name Suffix',
        'Organization', 'Job Title', 'Department',
        'Birthday',
        'Home Email', 'Work Email', 'Other Email',
        'Mobile Phone', 'Home Phone', 'Work Phone', 'Work Fax', 'Home Fax', 'Other Phone',
        'Home Address', 'Work Address', 'Other Address',
        'Notes', 'URLs'
    ]

    values = [headers]
    for contact in contacts:
        row = [contact.get(h, '') for h in headers]
        values.append(row)

    # Debug: Print what we're about to upload
    print(f"   Headers ({len(headers)}): {headers[:5]}...")
    print(f"   First row sample: {values[1][:5]}..." if len(values) > 1 else "No data rows")

    try:
        # Clear and upload
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Make sure we're sending the right data
        update_body = {
            'values': values,
            'majorDimension': 'ROWS'
        }

        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body=update_body
        ).execute()

        print(f"   Updated {result.get('updatedRows')} rows, {result.get('updatedColumns')} columns")

        # Format
        requests = [
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

        print(f"✅ Uploaded {len(contacts)} contacts with ALL {len(headers)} fields!")
        print(f"📋 View: https://docs.google.com/spreadsheets/d/{sheet_id}")
        return True

    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    print("=" * 60)
    print("TEST: 3 Contacts with ALL FIELDS")
    print("=" * 60)

    # Export 3 contacts
    contacts = export_3_all_fields()

    # Upload to sheets
    print("\n🔐 Authenticating...")
    service = authenticate_google_sheets()

    if upload_all_fields_to_sheets(service, contacts):
        print("\n🎉 SUCCESS!")
        print("✅ All fields exported and uploaded")
        print("✅ Check your Google Sheet for comprehensive data")
        print("\n🚀 Ready for full export of 3,792 contacts!")
    else:
        print("\n❌ Upload failed")

if __name__ == "__main__":
    main()