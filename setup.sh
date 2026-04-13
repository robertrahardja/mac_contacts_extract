#!/bin/bash

# Mac Contacts Export to Google Sheets - Setup Script
# This script automates the setup process

set -e

echo "🚀 Mac Contacts Export Setup Script"
echo "===================================="
echo ""

# Check if running on macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "❌ This script requires macOS to access Contacts"
    exit 1
fi

# Check Python version
echo "📝 Checking Python version..."
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3.8 or later."
    echo "   Visit: https://www.python.org/downloads/"
    exit 1
fi

PYTHON_VERSION=$(python3 -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
echo "✅ Python $PYTHON_VERSION found"

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv venv
    echo "✅ Virtual environment created"
else
    echo "✅ Virtual environment already exists"
fi

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source venv/bin/activate

# Upgrade pip
echo "📝 Upgrading pip..."
pip install --upgrade pip --quiet

# Install dependencies
echo "📦 Installing dependencies..."
pip install -r requirements.txt --quiet
echo "✅ All dependencies installed"

# Check for .env file
if [ ! -f ".env" ]; then
    echo ""
    echo "📋 Creating .env file from template..."
    cp .env.example .env
    echo "✅ .env file created"
    echo ""
    echo "⚠️  Please edit the .env file and add your Google Sheet ID"
    echo "   Open .env in your favorite editor and update GOOGLE_SHEET_ID"
else
    echo "✅ .env file already exists"
fi

# Check for Google API credentials
if [ ! -f "credentials.json" ]; then
    echo ""
    echo "🔑 Google API Setup Required"
    echo "=============================="
    echo ""
    echo "You need to set up Google API credentials. Follow these steps:"
    echo ""
    echo "1. Go to: https://console.cloud.google.com/"
    echo "2. Create a new project or select an existing one"
    echo "3. Enable the Google Sheets API"
    echo "4. Create credentials (OAuth 2.0 Client ID)"
    echo "5. Download the credentials as 'credentials.json'"
    echo "6. Place credentials.json in this directory"
    echo ""
    echo "For detailed instructions, see SETUP_GUIDE.md"
    echo ""
    read -p "Press Enter when you've placed credentials.json in this directory..."

    if [ ! -f "credentials.json" ]; then
        echo "❌ credentials.json not found. Please complete the setup and run this script again."
        exit 1
    fi
fi

echo "✅ Google API credentials found"

# Create exports directory
if [ ! -d "exports" ]; then
    mkdir -p exports
    echo "✅ Created exports directory for backups"
fi

echo ""
echo "✨ Setup Complete!"
echo "=================="
echo ""
echo "Next steps:"
echo "1. Make sure your .env file has your Google Sheet ID"
echo "2. Run the export with: ./run.sh"
echo "3. Or manually with: source venv/bin/activate && python mac_contacts_to_sheets.py"
echo ""
echo "For automation, see the README for cron job setup instructions."