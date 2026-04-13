#!/usr/bin/env python3
"""
Native Contacts Export using macOS Contacts Framework
Much faster than AppleScript
"""

import os
import sys
import json
import time
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# macOS Contacts framework
from Contacts import (
    CNContactStore,
    CNContactFetchRequest,
    CNContactGivenNameKey,
    CNContactFamilyNameKey,
    CNContactOrganizationNameKey,
    CNContactEmailAddressesKey,
    CNContactPhoneNumbersKey,
    CNContactPostalAddressesKey,
    CNContactNoteKey,
    CNContactMiddleNameKey,
    CNContactJobTitleKey,
    CNContactDepartmentNameKey,
    CNContactBirthdayKey,
    CNContactNicknameKey,
    CNContactNamePrefixKey,
    CNContactNameSuffixKey,
    CNContactUrlAddressesKey,
    CNLabelHome,
    CNLabelWork,
    CNLabelOther,
    CNLabelPhoneNumberMobile,
    CNLabelPhoneNumberiPhone,
    CNLabelPhoneNumberMain,
    CNLabelPhoneNumberHomeFax,
    CNLabelPhoneNumberWorkFax
)

load_dotenv()

# Google Sheets imports
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def export_contacts_native():
    """Export contacts using native macOS Contacts framework"""
    print("📦 Exporting all contacts using native macOS framework...")

    # Create contact store
    store = CNContactStore.alloc().init()

    # Request access to contacts
    # Note: This will prompt for permission if not already granted

    # Keys to fetch
    keys_to_fetch = [
        CNContactGivenNameKey,
        CNContactFamilyNameKey,
        CNContactMiddleNameKey,
        CNContactNicknameKey,
        CNContactNamePrefixKey,
        CNContactNameSuffixKey,
        CNContactOrganizationNameKey,
        CNContactJobTitleKey,
        CNContactDepartmentNameKey,
        CNContactEmailAddressesKey,
        CNContactPhoneNumbersKey,
        CNContactPostalAddressesKey,
        CNContactNoteKey,
        CNContactBirthdayKey,
        CNContactUrlAddressesKey
    ]

    # Create fetch request
    fetch_request = CNContactFetchRequest.alloc().initWithKeysToFetch_(keys_to_fetch)

    contacts_list = []
    error_ptr = None

    def process_contact(contact, stop):
        """Process each contact"""
        contact_dict = {}

        # Names
        contact_dict['First Name'] = contact.givenName() or ''
        contact_dict['Last Name'] = contact.familyName() or ''
        contact_dict['Middle Name'] = contact.middleName() or ''
        contact_dict['Nickname'] = contact.nickname() or ''
        contact_dict['Name Prefix'] = contact.namePrefix() or ''
        contact_dict['Name Suffix'] = contact.nameSuffix() or ''

        # Organization
        contact_dict['Organization'] = contact.organizationName() or ''
        contact_dict['Job Title'] = contact.jobTitle() or ''
        contact_dict['Department'] = contact.departmentName() or ''

        # Birthday
        if contact.birthday():
            bday = contact.birthday()
            contact_dict['Birthday'] = f"{bday.month}/{bday.day}/{bday.year}"
        else:
            contact_dict['Birthday'] = ''

        # Emails - numbered columns
        email_counts = {'home': 0, 'work': 0, 'other': 0}
        for email in contact.emailAddresses():
            label = str(email.label() or '')
            value = str(email.value())

            if CNLabelHome in label or 'home' in label.lower():
                email_counts['home'] += 1
                contact_dict[f"Home Email {email_counts['home']}"] = value
            elif CNLabelWork in label or 'work' in label.lower():
                email_counts['work'] += 1
                contact_dict[f"Work Email {email_counts['work']}"] = value
            else:
                email_counts['other'] += 1
                contact_dict[f"Other Email {email_counts['other']}"] = value

        # Phones - numbered columns
        phone_counts = {'mobile': 0, 'home': 0, 'work': 0, 'work_fax': 0, 'home_fax': 0, 'other': 0}
        for phone in contact.phoneNumbers():
            label = str(phone.label() or '')
            value = str(phone.value().stringValue())

            if CNLabelPhoneNumberMobile in label or CNLabelPhoneNumberiPhone in label or 'mobile' in label.lower():
                phone_counts['mobile'] += 1
                contact_dict[f"Mobile Phone {phone_counts['mobile']}"] = value
            elif CNLabelPhoneNumberWorkFax in label or ('work' in label.lower() and 'fax' in label.lower()):
                phone_counts['work_fax'] += 1
                contact_dict[f"Work Fax {phone_counts['work_fax']}"] = value
            elif CNLabelPhoneNumberHomeFax in label or ('home' in label.lower() and 'fax' in label.lower()):
                phone_counts['home_fax'] += 1
                contact_dict[f"Home Fax {phone_counts['home_fax']}"] = value
            elif CNLabelWork in label or CNLabelPhoneNumberMain in label or 'work' in label.lower():
                phone_counts['work'] += 1
                contact_dict[f"Work Phone {phone_counts['work']}"] = value
            elif CNLabelHome in label or 'home' in label.lower():
                phone_counts['home'] += 1
                contact_dict[f"Home Phone {phone_counts['home']}"] = value
            else:
                phone_counts['other'] += 1
                contact_dict[f"Other Phone {phone_counts['other']}"] = value

        # Addresses - numbered columns
        address_counts = {'home': 0, 'work': 0, 'other': 0}
        for address in contact.postalAddresses():
            label = str(address.label() or '')
            addr = address.value()

            addr_str = []
            if addr.street():
                addr_str.append(addr.street())
            if addr.city():
                addr_str.append(addr.city())
            if addr.state():
                addr_str.append(addr.state())
            if addr.postalCode():
                addr_str.append(addr.postalCode())
            if addr.country():
                addr_str.append(addr.country())

            full_address = ', '.join(addr_str)

            if CNLabelHome in label or 'home' in label.lower():
                address_counts['home'] += 1
                contact_dict[f"Home Address {address_counts['home']}"] = full_address
            elif CNLabelWork in label or 'work' in label.lower():
                address_counts['work'] += 1
                contact_dict[f"Work Address {address_counts['work']}"] = full_address
            else:
                address_counts['other'] += 1
                contact_dict[f"Other Address {address_counts['other']}"] = full_address

        # Notes
        contact_dict['Notes'] = contact.note() or ''

        # URLs - numbered columns
        url_count = 0
        for url in contact.urlAddresses():
            url_count += 1
            contact_dict[f"URL {url_count}"] = str(url.value())

        # Only add if contact has some data
        if contact_dict.get('First Name') or contact_dict.get('Last Name') or contact_dict.get('Organization'):
            contacts_list.append(contact_dict)

    # Enumerate contacts
    try:
        store.enumerateContactsWithFetchRequest_error_usingBlock_(
            fetch_request,
            error_ptr,
            process_contact
        )
        print(f"✅ Exported {len(contacts_list)} contacts")
        return contacts_list
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
    """Upload to Google Sheets with dynamic columns"""
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')

    if not contacts:
        print("❌ No contacts to upload")
        return False

    print(f"\n📊 Uploading {len(contacts)} contacts to Google Sheets...")

    # Collect all unique column names
    all_columns = set()
    for contact in contacts:
        all_columns.update(contact.keys())

    # Build headers
    base_headers = [
        'First Name', 'Last Name', 'Middle Name', 'Nickname',
        'Name Prefix', 'Name Suffix',
        'Organization', 'Job Title', 'Department',
        'Birthday'
    ]

    # Collect dynamic headers
    email_headers = sorted([col for col in all_columns if 'Email' in col])
    phone_headers = sorted([col for col in all_columns if 'Phone' in col or 'Fax' in col])
    address_headers = sorted([col for col in all_columns if 'Address' in col])
    url_headers = sorted([col for col in all_columns if col.startswith('URL')])

    headers = base_headers + email_headers + phone_headers + address_headers + ['Notes'] + url_headers

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

        # Upload all data
        print(f"   Uploading {len(contacts)} rows with {len(headers)} columns...")
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

def save_backup(contacts):
    """Save backup JSON"""
    if not contacts:
        return None

    export_dir = Path('exports')
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = export_dir / f"native_export_{timestamp}.json"

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts, f, indent=2, ensure_ascii=False)

    print(f"💾 Backup saved: {json_file}")
    return json_file

def main():
    print("="*60)
    print("NATIVE CONTACTS EXPORT - FULL")
    print("Using macOS Contacts Framework")
    print("="*60)

    start_time = time.time()

    # Export all contacts
    contacts = export_contacts_native()

    if not contacts:
        print("❌ No contacts exported")
        print("📝 Make sure you've granted permission to access Contacts")
        sys.exit(1)

    # Save backup
    save_backup(contacts)

    # Upload to Google Sheets
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    if upload_to_sheets(service, contacts):
        elapsed = time.time() - start_time
        print(f"\n🎉 EXPORT COMPLETE!")
        print(f"⏱️ Total time: {elapsed:.1f} seconds")
        print(f"📊 {len(contacts)} contacts exported with numbered columns")
    else:
        print("\n❌ Upload failed")

if __name__ == "__main__":
    main()