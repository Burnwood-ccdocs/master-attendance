#!/usr/bin/env python3
"""
Email automation module for sending attendance notifications to late and absent employees.
"""

import os
import pickle
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
from config import (
    EMAIL_AUTOMATION_ENABLED,
    EMAIL_SENDER,
    HR_EMAIL,
    GMAIL_OAUTH_CREDENTIALS,
    EMAIL_TEMPLATES
)

# Gmail API scope for sending emails
SCOPES = ['https://www.googleapis.com/auth/gmail.send']


class EmailAutomation:
    """Handles automated email notifications for attendance."""
    
    def __init__(self):
        self.service = None
        self.sender_email = EMAIL_SENDER
        self.hr_email = HR_EMAIL
        
        if EMAIL_AUTOMATION_ENABLED:
            self._authenticate()
    
    def _authenticate(self):
        """Authenticate with Gmail API using OAuth2."""
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
                # Create credentials from the config
                flow = InstalledAppFlow.from_client_config(
                    GMAIL_OAUTH_CREDENTIALS, SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(token_file, 'wb') as token:
                pickle.dump(creds, token)
        
        try:
            self.service = build('gmail', 'v1', credentials=creds)
            print("Email automation: Successfully authenticated with Gmail API")
        except Exception as e:
            print(f"Email automation: Failed to authenticate - {e}")
            self.service = None
    
    def _create_message(self, to, subject, body):
        """Create an email message."""
        message = MIMEMultipart()
        message['to'] = to
        message['from'] = self.sender_email
        message['subject'] = subject
        
        # Attach the body
        message.attach(MIMEText(body, 'plain'))
        
        # Encode the message
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        return {'raw': raw_message}
    
    def _send_message(self, message):
        """Send an email message."""
        if not self.service:
            print("Email automation: Service not initialized, skipping email")
            return False
        
        try:
            result = self.service.users().messages().send(
                userId='me', body=message).execute()
            print(f"Email sent successfully: {result['id']}")
            return True
        except HttpError as error:
            print(f'An error occurred while sending email: {error}')
            return False
    
    def send_late_notification(self, employee_data, date, start_time):
        """Send notification to a late employee."""
        if not EMAIL_AUTOMATION_ENABLED:
            return
        
        # Prepare template variables
        template_vars = {
            'name': employee_data['name'],
            'date': date.strftime('%Y-%m-%d'),
            'check_in_time': employee_data['arrival_time'],
            'expected_time': start_time.strftime('%I:%M %p'),
            'minutes_late': employee_data['minutes_late'],
            'hr_email': self.hr_email
        }
        
        # Get template
        template = EMAIL_TEMPLATES['late']
        subject = template['subject']
        body = template['body'].format(**template_vars)
        
        # Create and send message
        message = self._create_message(
            to=employee_data['email'],
            subject=subject,
            body=body
        )
        
        success = self._send_message(message)
        if success:
            print(f"  📧 Sent late notification to {employee_data['name']} ({employee_data['email']})")
        else:
            print(f"  ❌ Failed to send late notification to {employee_data['name']} ({employee_data['email']})")
        
        return success
    
    def send_absent_notification(self, employee_data, date):
        """Send notification to an absent employee."""
        if not EMAIL_AUTOMATION_ENABLED:
            return
        
        # Prepare template variables
        template_vars = {
            'name': employee_data['name'],
            'date': date.strftime('%Y-%m-%d'),
            'hr_email': self.hr_email
        }
        
        # Get template
        template = EMAIL_TEMPLATES['absent']
        subject = template['subject']
        body = template['body'].format(**template_vars)
        
        # Create and send message
        message = self._create_message(
            to=employee_data['email'],
            subject=subject,
            body=body
        )
        
        success = self._send_message(message)
        if success:
            print(f"  📧 Sent absent notification to {employee_data['name']} ({employee_data['email']})")
        else:
            print(f"  ❌ Failed to send absent notification to {employee_data['name']} ({employee_data['email']})")
        
        return success
    
    def send_batch_notifications(self, attendance_data, date, department_start_times):
        """Send batch notifications for all late and absent employees."""
        if not EMAIL_AUTOMATION_ENABLED:
            print("Email automation is disabled")
            return
        
        print("\n--- Starting Email Notifications ---")
        
        late_count = 0
        absent_count = 0
        
        for department, data in attendance_data.items():
            start_time = department_start_times.get(department)
            
            # Send notifications to late employees
            for employee in data.get('late', []):
                self.send_late_notification(employee, date, start_time)
                late_count += 1
            
            # Send notifications to absent employees
            for employee in data.get('absent', []):
                self.send_absent_notification(employee, date)
                absent_count += 1
        
        print(f"\n📧 Email Summary: Sent {late_count} late notifications, {absent_count} absent notifications")
        print("--- Email Notifications Complete ---\n")


# Utility function for one-time setup
def setup_gmail_oauth():
    """Run this to set up Gmail OAuth for the first time."""
    print("Setting up Gmail OAuth...")
    email_automation = EmailAutomation()
    if email_automation.service:
        print("✅ Gmail OAuth setup successful!")
        print("A token.pickle file has been created for future use.")
    else:
        print("❌ Gmail OAuth setup failed. Please check your credentials.")


if __name__ == "__main__":
    # Run setup when this file is executed directly
    setup_gmail_oauth() 