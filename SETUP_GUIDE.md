# Mac Contacts to Google Sheets - Setup Guide

## Overview
This solution automatically exports ALL your Mac contacts to Google Sheets without losing any data. It preserves:
- All names (first, last, middle, nicknames)
- All phone numbers with labels
- All email addresses with labels  
- All physical addresses
- Company and job information
- Birthdays and anniversaries
- Notes and custom fields
- Social media profiles
- Instant messaging accounts
- Related people
- URLs/websites
- Creation and modification dates

## Prerequisites

1. **Python 3** (comes pre-installed on modern Macs)
2. **Google Account** with access to Google Sheets
3. **Internet connection**

## Setup Instructions

### Step 1: Enable Google Sheets API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click "Select a project" → "NEW PROJECT"
3. Name it "Mac Contacts Export" and click "Create"
4. Once created, make sure it's selected
5. In the search bar, type "Google Sheets API"
6. Click on "Google Sheets API" and then "ENABLE"

### Step 2: Create OAuth 2.0 Credentials

1. In the left sidebar, click "Credentials"
2. Click "+ CREATE CREDENTIALS" → "OAuth client ID"
3. If prompted, configure the OAuth consent screen:
   - Choose "External" user type
   - Fill in required fields (app name, email)
   - Add your email to test users
   - Save and continue through all steps
4. Back in Credentials, click "+ CREATE CREDENTIALS" → "OAuth client ID" again
5. Choose "Desktop app" as the application type
6. Name it "Mac Contacts Exporter"
7. Click "Create"
8. Click "DOWNLOAD JSON" on the popup
9. Rename the downloaded file to `credentials.json`

### Step 3: Prepare the Script

1. Create a new folder on your Desktop called "ContactsExport"
2. Move the `credentials.json` file into this folder
3. Save the `mac_contacts_to_sheets.py` script to the same folder

### Step 4: Run the Export

1. Open Terminal (Applications → Utilities → Terminal)
2. Navigate to your folder:
   ```bash
   cd ~/Desktop/ContactsExport
   ```

3. Run the script:
   ```bash
   python3 mac_contacts_to_sheets.py
   ```

4. On first run:
   - Your browser will open asking you to authorize the app
   - Sign in with your Google account
   - Click "Continue" through any warnings (it's your own app)
   - Grant permissions for Google Sheets access
   - The browser will show "authentication successful"

5. The script will:
   - Extract all contacts from your Mac
   - Create a new Google Sheet
   - Upload all contact data
   - Format the sheet with filters and headers
   - Save a local JSON backup
   - Provide you with a direct link to the sheet

## What You Get

Your Google Sheet will have:
- **All contact fields** in separate columns
- **Formatted headers** in blue with white text
- **Auto-sized columns** for readability
- **Filters enabled** for easy searching/sorting
- **Multi-value fields** (like multiple emails) separated by semicolons
- **A local backup** in JSON format

## Scheduling Automatic Updates

To run this automatically every day/week:

1. Open Terminal and type:
   ```bash
   crontab -e
   ```

2. Add one of these lines:
   - For daily at 2 AM:
     ```
     0 2 * * * /usr/bin/python3 ~/Desktop/ContactsExport/mac_contacts_to_sheets.py
     ```
   - For weekly on Sundays at 2 AM:
     ```
     0 2 * * 0 /usr/bin/python3 ~/Desktop/ContactsExport/mac_contacts_to_sheets.py
     ```

3. Save and exit (Press Ctrl+X, then Y, then Enter)

## Troubleshooting

### "Permission denied" error
- Go to System Preferences → Security & Privacy → Privacy
- Select "Full Disk Access" or "Contacts"
- Add Terminal to the allowed apps

### "Module not found" error
The script will automatically install required packages, but if it fails:
```bash
pip3 install pyobjc-framework-AddressBook
pip3 install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

### Can't find credentials.json
Make sure you downloaded it from Google Cloud Console and placed it in the same folder as the script.

### Authentication keeps failing
Delete `token.pickle` file and try again:
```bash
rm token.pickle
```

## Privacy & Security

- Your contacts data never leaves your computer except to go directly to your own Google account
- The `credentials.json` file only allows access to Google Sheets in your account
- The `token.pickle` file stores your login session (keep it secure)
- Local backups are stored in JSON format on your computer

## Updating the Sheet

Each time you run the script, it creates a NEW sheet with a timestamp. This preserves history and prevents accidental data loss.

To update the same sheet instead of creating new ones, you can modify the script to use a specific spreadsheet ID.

## Support

If you encounter issues:
1. Check that your Mac Contacts app opens normally
2. Ensure you have internet connection
3. Verify your Google account has access to Google Sheets
4. Check the Terminal output for specific error messages
