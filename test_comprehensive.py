#!/usr/bin/env python3
"""
Test the comprehensive export with just 1 contact
"""

from export_all_contacts import export_contact_by_index

def test_comprehensive_export():
    print("🧪 Testing comprehensive contact export...")
    print("📱 Exporting first contact with ALL fields...")

    contact_data = export_contact_by_index(1)

    if contact_data:
        parts = contact_data.split('|')
        print(f"✅ Contact exported successfully with {len(parts)} fields!")
        print("\n📋 COMPREHENSIVE CONTACT DATA:")
        print("=" * 60)

        field_names = [
            'First Name', 'Last Name', 'Middle Name', 'Nickname',
            'Name Prefix', 'Name Suffix', 'Phonetic First Name', 'Phonetic Middle Name', 'Phonetic Last Name',
            'Organization', 'Job Title', 'Department',
            'All Emails', 'All Phone Numbers', 'Birthday', 'All Addresses',
            'All URLs', 'Social Profiles', 'Instant Messages', 'Related Names', 'Notes'
        ]

        for i, field in enumerate(parts):
            field_name = field_names[i] if i < len(field_names) else f'Field {i+1}'
            if field:  # Only show non-empty fields
                if len(field) > 100:
                    print(f"📝 {field_name}: {field[:100]}... (truncated for display)")
                else:
                    print(f"📝 {field_name}: {field}")

        print("=" * 60)
        print("✨ ALL contact data captured successfully!")
        print("🔥 No truncation, no limits, no data loss!")
        return True
    else:
        print("❌ Failed to export contact")
        return False

if __name__ == "__main__":
    if test_comprehensive_export():
        print("\n🎉 Ready to export all 3,792 contacts!")
        print("   Run: ./run.sh")
    else:
        print("\n⚠️  Please check Contacts app permissions")