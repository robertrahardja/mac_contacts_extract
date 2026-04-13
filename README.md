# Mac Contacts to Google Sheets - Complete Automation

## 🚀 Quick Start

This project is now fully set up with all dependencies installed!

### ✅ YOUR SETUP IS COMPLETE!

Everything is configured and ready:
- ✅ Google API credentials installed (`credentials.json`)
- ✅ Google Sheet ID configured: `1Kqa9G-tPp6AlLaaUV3Of3KoEhXpm74PGtZd7yrYNwDk`
- ✅ Python dependencies installed in `venv`
- ✅ Environment variables configured in `.env`

### 🎯 Run Your Export - NO TIMEOUT ISSUES!

**✨ NEW: Full export script that handles all 3,792 contacts without timeout!**

```bash
# Export ALL contacts directly to Google Sheets
./run.sh
```

This will:
- Process all 3,792 contacts individually (no timeouts!)
- Show a progress bar as it exports
- Save local backups (JSON and CSV)
- Upload everything to your Google Sheet
- Expected time: 5-10 minutes

### 🔥 COMPREHENSIVE EXPORT - NO DATA LOSS!

**✨ EVERY SINGLE FIELD** from your Contacts app is exported:

- **Names**: First, Middle, Last, Nickname, Prefix, Suffix
- **Phonetic Names**: Phonetic First, Middle, Last
- **Organization**: Company, Job Title, Department
- **Contact Info**: ALL emails (with labels), ALL phone numbers (with labels)
- **Addresses**: ALL addresses (with labels, street, city, state, zip, country)
- **Online Presence**: ALL URLs, ALL social profiles, ALL instant messaging accounts
- **Personal**: Birthday, ALL related names (family members)
- **Notes**: Complete notes (NO TRUNCATION)

**Nothing is truncated, limited, or lost!**

### Alternative Options:

1. **Quick test** (verify setup works):
   ```bash
   source venv/bin/activate
   python export_first_100.py
   ```

2. **Manual export** (if you prefer CSV):
   ```bash
   osascript export_contacts.applescript
   ```

---

## Two Solutions - Choose What Works Best for You

### Solution 1: Direct Python Export (Fully Automated) ✅ READY TO USE
**Best for:** Complete automation, preserves ALL data fields, creates formatted Google Sheets automatically
- ✅ Exports directly to Google Sheets (no manual import needed)
- ✅ Preserves 20+ contact fields including social media, IMs, related people
- ✅ Automatically formats the sheet with filters and styling
- ✅ Can be scheduled to run automatically
- ✅ Creates timestamped backups
- ✅ All dependencies already installed in venv
- ⚠️ Requires one-time Google API setup (10 minutes)

### Solution 2: AppleScript Export to CSV
**Best for:** Quick export, no API setup needed, manual import to Google Sheets
- ✅ No API setup required
- ✅ Simple double-click to run
- ✅ Exports main contact fields
- ✅ Creates CSV file ready for Google Sheets import
- ⚠️ Requires manual upload to Google Sheets
- ⚠️ May not capture all custom fields

---

## 📋 Setup Instructions

### Step 1: Configure Environment
```bash
# Copy the example environment file
cp .env.example .env

# Edit .env and add your Google Sheet ID
# Find this in your Google Sheet URL: docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit
```

### Step 2: Google API Setup (One-time)

1. **Enable Google Sheets API:**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing
   - Enable "Google Sheets API"

2. **Create Credentials:**
   - Go to Credentials → Create Credentials → OAuth 2.0 Client ID
   - Application type: Desktop app
   - Download as `credentials.json`
   - Place in project root directory

3. **First Run Authentication:**
   ```bash
   ./run.sh
   # Browser will open for authorization
   # Grant permissions to access Google Sheets
   ```

### Step 3: Run the Export

#### Using the provided scripts:
```bash
# Setup (first time only)
./setup.sh

# Run export
./run.sh
```

#### Or manually:
```bash
# Activate virtual environment
source venv/bin/activate

# Run the export
python mac_contacts_to_sheets.py
```

### Option 2: AppleScript Quick Export

1. **Save and run the AppleScript:**
   - Save `export_contacts.applescript` to your Desktop
   - Double-click to open in Script Editor
   - Click the "Run" button (▶️)

2. **What happens:**
   - All contacts exported to CSV on Desktop
   - Dialog asks if you want to open Google Sheets
   - Instructions provided for importing

3. **Import to Google Sheets:**
   - Go to [sheets.google.com](https://sheets.google.com)
   - Create new spreadsheet
   - File → Import → Upload your CSV
   - Choose "Replace current sheet"
   - Click "Import data"

---

## Comparison Table

| Feature | Python Solution | AppleScript Solution |
|---------|----------------|---------------------|
| **Setup Time** | 10 minutes (first time) | None |
| **Automation** | Fully automated | Semi-automated |
| **Google Sheets** | Direct upload | Manual import |
| **Fields Exported** | ALL (20+) | Main fields (15) |
| **Formatting** | Automatic | Manual |
| **Scheduling** | Yes (cron) | Yes (Calendar) |
| **Backups** | JSON + Sheet | CSV only |
| **Updates** | Creates new sheets | Overwrites CSV |

---

## Data Preserved

### Python Solution Exports Everything:
- Names (first, last, middle, nickname, maiden)
- Phone numbers (all types with labels)
- Email addresses (all with labels)
- Physical addresses (all with full details)
- Organization (company, title, department)
- Dates (birthday, anniversary)
- Instant messaging (AIM, Jabber, MSN, Yahoo, ICQ)
- Social profiles (Facebook, Twitter, LinkedIn, etc.)
- URLs/websites
- Notes (full text)
- Related people (spouse, child, parent, etc.)
- Custom fields
- Contact ID
- Creation date
- Modification date

### AppleScript Exports Main Fields:
- Names (first, last, middle, nickname)
- Phone numbers (up to 4)
- Email addresses (up to 3)
- Addresses (home and work)
- Organization info
- Birthday
- Notes
- URLs
- Social media profiles

---

## Automation Options

### Schedule Daily/Weekly Updates (Python)

Add to crontab for automatic updates:
```bash
# Daily at 2 AM
0 2 * * * /usr/bin/python3 ~/Desktop/mac_contacts_to_sheets.py

# Weekly on Sundays
0 2 * * 0 /usr/bin/python3 ~/Desktop/mac_contacts_to_sheets.py

# Monthly on the 1st
0 2 1 * * /usr/bin/python3 ~/Desktop/mac_contacts_to_sheets.py
```

### Schedule with Calendar (AppleScript)

1. Open Calendar app
2. Create recurring event
3. Set alert to "Open file"
4. Select the AppleScript file
5. Set to run at event time

---

## Troubleshooting

### Common Issues and Fixes

**"Permission denied" accessing contacts:**
- System Preferences → Security & Privacy → Privacy → Contacts
- Add Terminal (for Python) or Script Editor (for AppleScript)

**Python modules not installing:**
```bash
pip3 install --user pyobjc-framework-AddressBook
pip3 install --user google-auth google-auth-oauthlib google-api-python-client
```

**AppleScript timeout with many contacts:**
- Split export into groups in Contacts app
- Export each group separately

**Google authentication fails:**
- Delete `token.pickle` and try again
- Check that Google Sheets API is enabled
- Ensure `credentials.json` is in the right place

---

## Security & Privacy

- **Your data stays private** - Goes directly from your Mac to your Google account
- **No third-party access** - You control all credentials
- **Local backups** - Keep copies on your Mac
- **Revokable access** - Can remove app access anytime in Google settings

---

## Advanced Features

### Modify What Gets Exported

Edit either script to customize fields:
- Comment out unwanted fields
- Add custom field mappings
- Filter contacts by group
- Export only recent changes

### Integration Ideas

- Sync with CRM systems
- Create mailing lists
- Birthday reminders
- Contact deduplication
- Business card scanning

---

## Support

Before asking for help, check:
1. ✓ Contacts app opens normally
2. ✓ Internet connection works
3. ✓ Google account has Sheets access
4. ✓ Terminal/Script Editor has Contacts permission
5. ✓ Error messages in output

---

## Files Included

1. **mac_contacts_to_sheets.py** - Python automation script
2. **export_contacts.applescript** - AppleScript CSV exporter  
3. **SETUP_GUIDE.md** - Detailed setup instructions
4. **README.md** - This file

Choose the solution that fits your needs and comfort level. The Python solution is more powerful but requires initial setup. The AppleScript is simpler but requires manual import to Google Sheets.
