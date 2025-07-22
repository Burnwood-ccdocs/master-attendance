#!/usr/bin/env python3
"""
Setup script for Gmail OAuth authentication.
Run this script once to authenticate and create the token.pickle file.
"""

from email_automation import setup_gmail_oauth
import os
from config import EMAIL_SENDER

def main():
    print("=== Gmail OAuth Setup for Attendance Email Automation ===")
    print(f"\nThis will set up email sending from: {EMAIL_SENDER}")
    print("\nIMPORTANT: Make sure you:")
    print("1. Have enabled Gmail API in Google Cloud Console")
    print("2. Are logged into the Google account that will send emails")
    print("3. Have the correct OAuth credentials in config.py")
    
    input("\nPress Enter to continue with setup...")
    
    # Run the OAuth setup
    setup_gmail_oauth()
    
    # Check if token was created
    if os.path.exists('token.pickle'):
        print("\n✅ Setup complete! Token file created.")
        print("Email automation is now ready to use.")
        print("\nNote: The system will send emails automatically during EOD reports to:")
        print("- Late employees (checked in 6-30 minutes after start time)")
        print("- Absent employees (no check-in or >30 minutes late)")
    else:
        print("\n❌ Setup failed. Please check the error messages above.")
        print("Common issues:")
        print("1. Invalid OAuth credentials")
        print("2. Gmail API not enabled")
        print("3. Network connectivity issues")

if __name__ == "__main__":
    main() 