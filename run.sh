#!/bin/bash

# Mac Contacts Export to Google Sheets - Run Script
# Simple script to run the export

set -e

echo "🚀 Starting Mac Contacts Export to Google Sheets"
echo "================================================"
echo ""

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "❌ Virtual environment not found. Please run ./setup.sh first"
    exit 1
fi

# Check if .env file exists
if [ ! -f ".env" ]; then
    echo "❌ .env file not found. Please run ./setup.sh first"
    exit 1
fi

# Check if credentials.json exists
if [ ! -f "credentials.json" ]; then
    echo "❌ credentials.json not found. Please complete Google API setup"
    echo "   See SETUP_GUIDE.md for instructions"
    exit 1
fi

# Activate virtual environment
source venv/bin/activate

# Run the no-timeout export that handles all contacts
echo "📊 Exporting ALL contacts to Google Sheets..."
echo ""
echo "NOTE: If prompted, grant Terminal access to Contacts in:"
echo "      System Settings > Privacy & Security > Contacts"
echo ""
echo "This will process all your contacts individually to avoid timeouts."
echo "Expected time: 5-10 minutes for 3,792 contacts"
echo ""

python export_comprehensive_stable.py

echo ""
echo "✨ Export complete!"