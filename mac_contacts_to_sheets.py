#!/usr/bin/env python3
"""
Mac Contacts to Google Sheets Exporter
Automatically exports all Mac contacts to Google Sheets without data loss
"""

import os
import sys
import json
import pickle
from datetime import datetime
from pathlib import Path

# Mac-specific imports - using modern Contacts framework
try:
    import Contacts
    import Foundation
    import objc
except ImportError:
    print("Installing required macOS frameworks...")
    os.system("pip3 install pyobjc-framework-Contacts pyobjc-framework-Cocoa")
    import Contacts
    import Foundation
    import objc

# Google Sheets imports
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
except ImportError:
    print("Installing required Google libraries...")
    os.system("pip3 install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_mac_contacts():
    """Extract all contacts from Mac Address Book with all fields"""
    
    # Access the Address Book
    address_book = ABAddressBook.sharedAddressBook()
    all_people = address_book.people()
    
    contacts_data = []
    
    for person in all_people:
        contact = {}
        
        # Basic Information
        contact['First Name'] = person.valueForProperty_(kABFirstNameProperty) or ''
        contact['Last Name'] = person.valueForProperty_(kABLastNameProperty) or ''
        contact['Middle Name'] = person.valueForProperty_(kABMiddleNameProperty) or ''
        contact['Nickname'] = person.valueForProperty_(kABNicknameProperty) or ''
        
        # Organization
        contact['Company'] = person.valueForProperty_(kABOrganizationProperty) or ''
        contact['Job Title'] = person.valueForProperty_(kABJobTitleProperty) or ''
        contact['Department'] = person.valueForProperty_(kABDepartmentProperty) or ''
        
        # Birthday
        birthday = person.valueForProperty_(kABBirthdayProperty)
        if birthday:
            contact['Birthday'] = birthday.description()
        else:
            contact['Birthday'] = ''
        
        # Notes
        contact['Notes'] = person.valueForProperty_(kABNoteProperty) or ''
        
        # Email addresses (multi-value)
        emails = person.valueForProperty_(kABEmailProperty)
        if emails:
            email_list = []
            for i in range(emails.count()):
                email = emails.valueAtIndex_(i)
                label = emails.labelAtIndex_(i)
                email_list.append(f"{label}: {email}")
            contact['Emails'] = '; '.join(email_list)
        else:
            contact['Emails'] = ''
        
        # Phone numbers (multi-value)
        phones = person.valueForProperty_(kABPhoneProperty)
        if phones:
            phone_list = []
            for i in range(phones.count()):
                phone = phones.valueAtIndex_(i)
                label = phones.labelAtIndex_(i)
                phone_list.append(f"{label}: {phone}")
            contact['Phone Numbers'] = '; '.join(phone_list)
        else:
            contact['Phone Numbers'] = ''
        
        # Addresses (multi-value)
        addresses = person.valueForProperty_(kABAddressProperty)
        if addresses:
            address_list = []
            for i in range(addresses.count()):
                address = addresses.valueAtIndex_(i)
                label = addresses.labelAtIndex_(i)
                
                # Extract address components
                street = address.get('Street', '')
                city = address.get('City', '')
                state = address.get('State', '')
                zip_code = address.get('ZIP', '')
                country = address.get('Country', '')
                
                full_address = f"{label}: {street}, {city}, {state} {zip_code}, {country}"
                address_list.append(full_address.strip(', '))
            contact['Addresses'] = '; '.join(address_list)
        else:
            contact['Addresses'] = ''
        
        # URLs (multi-value)
        urls = person.valueForProperty_(kABURLsProperty)
        if urls:
            url_list = []
            for i in range(urls.count()):
                url = urls.valueAtIndex_(i)
                label = urls.labelAtIndex_(i)
                url_list.append(f"{label}: {url}")
            contact['URLs'] = '; '.join(url_list)
        else:
            contact['URLs'] = ''
        
        # Instant Message (multi-value)
        ims = person.valueForProperty_(kABInstantMessageProperty)
        if ims:
            im_list = []
            for i in range(ims.count()):
                im = ims.valueAtIndex_(i)
                label = ims.labelAtIndex_(i)
                service = im.get('Service', '')
                username = im.get('Username', '')
                im_list.append(f"{label} ({service}): {username}")
            contact['Instant Messages'] = '; '.join(im_list)
        else:
            contact['Instant Messages'] = ''
        
        # Social Profiles (multi-value)
        social = person.valueForProperty_(kABSocialProfileProperty)
        if social:
            social_list = []
            for i in range(social.count()):
                profile = social.valueAtIndex_(i)
                label = social.labelAtIndex_(i)
                service = profile.get('service', '')
                username = profile.get('username', '')
                url = profile.get('url', '')
                social_list.append(f"{service}: {username} ({url})")
            contact['Social Profiles'] = '; '.join(social_list)
        else:
            contact['Social Profiles'] = ''
        
        # Related Names (multi-value)
        related = person.valueForProperty_(kABRelatedNamesProperty)
        if related:
            related_list = []
            for i in range(related.count()):
                name = related.valueAtIndex_(i)
                label = related.labelAtIndex_(i)
                related_list.append(f"{label}: {name}")
            contact['Related Names'] = '; '.join(related_list)
        else:
            contact['Related Names'] = ''
        
        # Unique ID (for reference)
        contact['Contact ID'] = person.uniqueId()
        
        # Creation and modification dates
        creation_date = person.valueForProperty_('kABCreationDateProperty')
        if creation_date:
            contact['Created Date'] = str(creation_date)
        else:
            contact['Created Date'] = ''
            
        modification_date = person.valueForProperty_('kABModificationDateProperty')
        if modification_date:
            contact['Modified Date'] = str(modification_date)
        else:
            contact['Modified Date'] = ''
        
        contacts_data.append(contact)
    
    return contacts_data

def authenticate_google_sheets():
    """Authenticate and return Google Sheets service"""
    creds = None
    token_file = 'token.pickle'
    
    # Token file stores the user's access and refresh tokens
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # You'll need to create credentials.json from Google Cloud Console
            if not os.path.exists('credentials.json'):
                print("\n❌ credentials.json file not found!")
                print("Please follow these steps:")
                print("1. Go to https://console.cloud.google.com/")
                print("2. Create a new project or select existing")
                print("3. Enable Google Sheets API")
                print("4. Create credentials (OAuth 2.0 Client ID)")
                print("5. Download the credentials.json file")
                print("6. Place it in the same directory as this script")
                sys.exit(1)
                
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)
    
    return build('sheets', 'v4', credentials=creds)

def create_or_update_sheet(service, contacts_data):
    """Create a new Google Sheet or update existing one with contacts data"""
    
    # Create a new spreadsheet
    spreadsheet = {
        'properties': {
            'title': f'Mac Contacts Export - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        }
    }
    
    spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
    spreadsheet_id = spreadsheet.get('spreadsheetId')
    
    print(f"✅ Created new spreadsheet with ID: {spreadsheet_id}")
    print(f"📊 URL: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
    
    # Prepare data for Google Sheets
    if not contacts_data:
        print("No contacts found to export.")
        return
    
    # Get all unique headers from all contacts
    all_headers = set()
    for contact in contacts_data:
        all_headers.update(contact.keys())
    headers = sorted(list(all_headers))
    
    # Create values array with headers
    values = [headers]
    
    # Add contact data
    for contact in contacts_data:
        row = [contact.get(header, '') for header in headers]
        values.append(row)
    
    # Update the sheet with data
    body = {
        'values': values
    }
    
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range='A1',
        valueInputOption='RAW',
        body=body
    ).execute()
    
    print(f"✅ Successfully exported {len(contacts_data)} contacts")
    print(f"📝 Updated {result.get('updatedCells')} cells")
    
    # Format the header row
    requests = [
        {
            'repeatCell': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': 0,
                    'endRowIndex': 1
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.2,
                            'green': 0.5,
                            'blue': 0.8
                        },
                        'textFormat': {
                            'foregroundColor': {
                                'red': 1.0,
                                'green': 1.0,
                                'blue': 1.0
                            },
                            'bold': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat)'
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
        },
        {
            'setBasicFilter': {
                'filter': {
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': 0,
                        'startColumnIndex': 0,
                        'endColumnIndex': len(headers)
                    }
                }
            }
        }
    ]
    
    body = {
        'requests': requests
    }
    
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    
    print("✅ Applied formatting and filters")
    
    return spreadsheet_id

def main():
    """Main function to run the export"""
    print("🚀 Mac Contacts to Google Sheets Exporter")
    print("=" * 50)
    
    # Step 1: Extract contacts from Mac
    print("\n📱 Extracting contacts from Mac Address Book...")
    contacts = get_mac_contacts()
    print(f"✅ Found {len(contacts)} contacts")
    
    # Step 2: Authenticate with Google
    print("\n🔐 Authenticating with Google Sheets...")
    service = authenticate_google_sheets()
    print("✅ Authentication successful")
    
    # Step 3: Create and populate Google Sheet
    print("\n📤 Uploading contacts to Google Sheets...")
    spreadsheet_id = create_or_update_sheet(service, contacts)
    
    # Step 4: Save a local backup
    backup_file = f"contacts_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(backup_file, 'w') as f:
        json.dump(contacts, f, indent=2, default=str)
    print(f"\n💾 Local backup saved to: {backup_file}")
    
    print("\n✨ Export complete!")
    print(f"🔗 Open your spreadsheet: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")

if __name__ == "__main__":
    main()
