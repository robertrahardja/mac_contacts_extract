#!/usr/bin/env python3
"""
Mac Contacts to Google Sheets Exporter (Modern Version)
Uses the modern Contacts framework for macOS
"""

import os
import sys
import json
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Mac-specific imports - using modern Contacts framework
import Contacts
import Foundation
import objc

# Google Sheets imports
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

# Google Sheets API scope
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_contacts():
    """Fetch all contacts from Mac Contacts app using modern framework"""
    print("📱 Fetching contacts from Mac Contacts app...")

    # Create contact store
    store = Contacts.CNContactStore.alloc().init()

    # Request access to contacts
    # Check if we already have access
    auth_status = Contacts.CNContactStore.authorizationStatusForEntityType_(Contacts.CNEntityTypeContacts)

    if auth_status != Contacts.CNAuthorizationStatusAuthorized:
        # Request access synchronously
        granted = store.requestAccessForEntityType_(Contacts.CNEntityTypeContacts)
        if not granted:
            print("❌ Access to Contacts denied. Please grant permission in:")
            print("   System Settings > Privacy & Security > Contacts")
            print("   Then run this script again.")
            sys.exit(1)

    # Define which properties to fetch
    keys_to_fetch = [
        Contacts.CNContactGivenNameKey,
        Contacts.CNContactFamilyNameKey,
        Contacts.CNContactMiddleNameKey,
        Contacts.CNContactNicknameKey,
        Contacts.CNContactOrganizationNameKey,
        Contacts.CNContactJobTitleKey,
        Contacts.CNContactDepartmentNameKey,
        Contacts.CNContactEmailAddressesKey,
        Contacts.CNContactPhoneNumbersKey,
        Contacts.CNContactPostalAddressesKey,
        Contacts.CNContactBirthdayKey,
        Contacts.CNContactNoteKey,
        Contacts.CNContactUrlAddressesKey,
        Contacts.CNContactInstantMessageAddressesKey,
        Contacts.CNContactSocialProfilesKey,
        Contacts.CNContactRelationsKey,
        Contacts.CNContactIdentifierKey,
    ]

    # Create fetch request
    request = Contacts.CNContactFetchRequest.alloc().initWithKeysToFetch_(keys_to_fetch)

    contacts_data = []

    def process_contact(contact, stop):
        """Process each contact"""
        contact_dict = {}

        # Names
        contact_dict['First Name'] = contact.givenName() or ''
        contact_dict['Last Name'] = contact.familyName() or ''
        contact_dict['Middle Name'] = contact.middleName() or ''
        contact_dict['Nickname'] = contact.nickname() or ''

        # Organization
        contact_dict['Organization'] = contact.organizationName() or ''
        contact_dict['Job Title'] = contact.jobTitle() or ''
        contact_dict['Department'] = contact.departmentName() or ''

        # Birthday
        birthday = contact.birthday()
        if birthday:
            contact_dict['Birthday'] = f"{birthday.month}/{birthday.day}/{birthday.year if birthday.year else ''}"
        else:
            contact_dict['Birthday'] = ''

        # Note
        contact_dict['Note'] = contact.note() or ''

        # Email addresses
        emails = contact.emailAddresses()
        if emails and emails.count() > 0:
            email_list = []
            for i in range(emails.count()):
                labeled_value = emails.objectAtIndex_(i)
                label = Contacts.CNLabeledValue.localizedStringForLabel_(labeled_value.label()) if labeled_value.label() else 'Other'
                email = labeled_value.value()
                email_list.append(f"{label}: {email}")
            contact_dict['Email Addresses'] = '; '.join(email_list)
        else:
            contact_dict['Email Addresses'] = ''

        # Phone numbers
        phones = contact.phoneNumbers()
        if phones and phones.count() > 0:
            phone_list = []
            for i in range(phones.count()):
                labeled_value = phones.objectAtIndex_(i)
                label = Contacts.CNLabeledValue.localizedStringForLabel_(labeled_value.label()) if labeled_value.label() else 'Other'
                phone = labeled_value.value().stringValue()
                phone_list.append(f"{label}: {phone}")
            contact_dict['Phone Numbers'] = '; '.join(phone_list)
        else:
            contact_dict['Phone Numbers'] = ''

        # Postal addresses
        addresses = contact.postalAddresses()
        if addresses and addresses.count() > 0:
            address_list = []
            for i in range(addresses.count()):
                labeled_value = addresses.objectAtIndex_(i)
                label = Contacts.CNLabeledValue.localizedStringForLabel_(labeled_value.label()) if labeled_value.label() else 'Other'
                address = labeled_value.value()

                street = address.street() or ''
                city = address.city() or ''
                state = address.state() or ''
                postal_code = address.postalCode() or ''
                country = address.country() or ''

                full_address = f"{street}, {city}, {state} {postal_code}, {country}".strip(', ')
                address_list.append(f"{label}: {full_address}")
            contact_dict['Addresses'] = '; '.join(address_list)
        else:
            contact_dict['Addresses'] = ''

        # URLs
        urls = contact.urlAddresses()
        if urls and urls.count() > 0:
            url_list = []
            for i in range(urls.count()):
                labeled_value = urls.objectAtIndex_(i)
                label = Contacts.CNLabeledValue.localizedStringForLabel_(labeled_value.label()) if labeled_value.label() else 'Other'
                url = labeled_value.value()
                url_list.append(f"{label}: {url}")
            contact_dict['URLs'] = '; '.join(url_list)
        else:
            contact_dict['URLs'] = ''

        # Instant messages
        ims = contact.instantMessageAddresses()
        if ims and ims.count() > 0:
            im_list = []
            for i in range(ims.count()):
                labeled_value = ims.objectAtIndex_(i)
                label = Contacts.CNLabeledValue.localizedStringForLabel_(labeled_value.label()) if labeled_value.label() else 'Other'
                im = labeled_value.value()
                service = im.service() or ''
                username = im.username() or ''
                im_list.append(f"{label} ({service}): {username}")
            contact_dict['Instant Messages'] = '; '.join(im_list)
        else:
            contact_dict['Instant Messages'] = ''

        # Social profiles
        social = contact.socialProfiles()
        if social and social.count() > 0:
            social_list = []
            for i in range(social.count()):
                labeled_value = social.objectAtIndex_(i)
                profile = labeled_value.value()
                service = profile.service() or ''
                username = profile.username() or ''
                url = profile.urlString() or ''
                social_list.append(f"{service}: {username} ({url})")
            contact_dict['Social Profiles'] = '; '.join(social_list)
        else:
            contact_dict['Social Profiles'] = ''

        # Relations
        relations = contact.contactRelations()
        if relations and relations.count() > 0:
            relation_list = []
            for i in range(relations.count()):
                labeled_value = relations.objectAtIndex_(i)
                label = Contacts.CNLabeledValue.localizedStringForLabel_(labeled_value.label()) if labeled_value.label() else 'Other'
                relation = labeled_value.value()
                name = relation.name() or ''
                relation_list.append(f"{label}: {name}")
            contact_dict['Related Names'] = '; '.join(relation_list)
        else:
            contact_dict['Related Names'] = ''

        # Contact ID
        contact_dict['Contact ID'] = contact.identifier()

        contacts_data.append(contact_dict)

    # Fetch all contacts
    error_ptr = objc.nil
    success = store.enumerateContactsWithFetchRequest_error_usingBlock_(
        request,
        None,  # Error handling not needed for enumeration
        process_contact
    )

    if not success:
        print("❌ Error fetching contacts")
        sys.exit(1)

    print(f"✅ Found {len(contacts_data)} contacts")
    return contacts_data

def authenticate_google_sheets():
    """Authenticate and return Google Sheets service"""
    creds = None
    token_file = 'token.json'

    # Token file stores the user's access and refresh tokens
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("❌ credentials.json not found!")
                print("Please follow the setup guide to get your Google API credentials")
                sys.exit(1)

            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for next run
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service

def create_or_update_sheet(service, contacts_data):
    """Create or update Google Sheet with contacts data"""

    # Get sheet ID from environment
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("❌ GOOGLE_SHEET_ID not found in .env file")
        sys.exit(1)

    sheet_name = os.getenv('SHEET_NAME', 'Contacts')

    print(f"📊 Updating Google Sheet: {sheet_id}")

    # Prepare headers
    if contacts_data:
        headers = list(contacts_data[0].keys())
    else:
        print("No contacts found to export")
        return None

    # Prepare values (headers + data rows)
    values = [headers]
    for contact in contacts_data:
        row = [contact.get(header, '') for header in headers]
        values.append(row)

    try:
        # Clear existing content
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:Z"
        ).execute()

        # Update with new data
        body = {
            'values': values
        }

        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body=body
        ).execute()

        # Format the sheet
        requests = [
            # Freeze header row
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': 0,
                        'gridProperties': {
                            'frozenRowCount': 1
                        }
                    },
                    'fields': 'gridProperties.frozenRowCount'
                }
            },
            # Bold header row
            {
                'repeatCell': {
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': 0,
                        'endRowIndex': 1
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.textFormat.bold'
                }
            },
            # Auto-resize columns
            {
                'autoResizeDimensions': {
                    'dimensions': {
                        'sheetId': 0,
                        'dimension': 'COLUMNS',
                        'startIndex': 0,
                        'endIndex': len(headers)
                    }
                }
            },
            # Add filter
            {
                'setBasicFilter': {
                    'filter': {
                        'range': {
                            'sheetId': 0,
                            'startRowIndex': 0,
                            'endRowIndex': len(values),
                            'startColumnIndex': 0,
                            'endColumnIndex': len(headers)
                        }
                    }
                }
            }
        ]

        batch_update_body = {'requests': requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body=batch_update_body
        ).execute()

        print(f"✅ Updated {result.get('updatedCells')} cells")
        print(f"📋 Sheet URL: https://docs.google.com/spreadsheets/d/{sheet_id}")

        return sheet_id

    except Exception as e:
        print(f"❌ Error updating sheet: {str(e)}")
        print("\nMake sure:")
        print("1. The Google Sheet ID in .env is correct")
        print("2. The sheet exists and you have edit access")
        print("3. The Google Sheets API is enabled in your project")
        return None

def save_local_backup(contacts_data):
    """Save local backup of contacts"""
    export_dir = Path(os.getenv('EXPORT_DIR', 'exports'))
    export_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Save as JSON
    json_file = export_dir / f"contacts_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(contacts_data, f, indent=2, ensure_ascii=False)

    print(f"💾 Local backup saved: {json_file}")
    return json_file

def main():
    """Main function"""
    print("=" * 50)
    print("Mac Contacts to Google Sheets Exporter")
    print("=" * 50)

    # Get contacts from Mac
    contacts = get_contacts()

    if not contacts:
        print("No contacts found in Contacts app")
        return

    # Save local backup
    backup_file = save_local_backup(contacts)

    # Authenticate with Google
    print("\n🔐 Authenticating with Google...")
    service = authenticate_google_sheets()

    # Update Google Sheet
    sheet_id = create_or_update_sheet(service, contacts)

    if sheet_id:
        print("\n✨ Export complete!")
        print(f"📱 Total contacts exported: {len(contacts)}")
        print(f"💾 Backup saved to: {backup_file}")
        print(f"📊 View your sheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
    else:
        print("\n⚠️ Export to Google Sheets failed, but local backup was saved")

if __name__ == "__main__":
    main()