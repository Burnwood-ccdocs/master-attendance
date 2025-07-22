#!/usr/bin/env python3
"""
Test script for email automation functionality.
"""

import datetime
import pytz
from email_automation import EmailAutomation
from config import TIMEZONE, EMAIL_SENDER, HR_EMAIL

def test_email_templates():
    """Test email template formatting."""
    print("=== Testing Email Templates ===")
    
    # Test data
    test_employee_late = {
        'name': 'Test Employee',
        'email': 'garbhits@ccdocs.com',
        'arrival_time': '09:15 AM',
        'minutes_late': 15
    }
    
    test_employee_absent = {
        'name': 'Test Employee 2',
        'email': 'test2@example.com'
    }
    
    date = datetime.datetime.now(pytz.timezone(TIMEZONE)).date()
    start_time = datetime.datetime.strptime("09:00", "%H:%M").time()
    start_dt = datetime.datetime.combine(date, start_time, tzinfo=pytz.timezone(TIMEZONE))
    
    # Create email automation instance
    email_automation = EmailAutomation()
    
    print("\n1. Testing Late Employee Email:")
    print("-" * 50)
    
    # Create late notification (but don't send)
    from config import EMAIL_TEMPLATES
    template_vars = {
        'name': test_employee_late['name'],
        'date': date.strftime('%Y-%m-%d'),
        'check_in_time': test_employee_late['arrival_time'],
        'expected_time': start_dt.strftime('%I:%M %p'),
        'minutes_late': test_employee_late['minutes_late'],
        'hr_email': HR_EMAIL
    }
    
    late_body = EMAIL_TEMPLATES['late']['body'].format(**template_vars)
    print(f"Subject: {EMAIL_TEMPLATES['late']['subject']}")
    print(f"To: {test_employee_late['email']}")
    print(f"From: {EMAIL_SENDER}")
    print("\nBody:")
    print(late_body)
    
    print("\n2. Testing Absent Employee Email:")
    print("-" * 50)
    
    template_vars = {
        'name': test_employee_absent['name'],
        'date': date.strftime('%Y-%m-%d'),
        'hr_email': HR_EMAIL
    }
    
    absent_body = EMAIL_TEMPLATES['absent']['body'].format(**template_vars)
    print(f"Subject: {EMAIL_TEMPLATES['absent']['subject']}")
    print(f"To: {test_employee_absent['email']}")
    print(f"From: {EMAIL_SENDER}")
    print("\nBody:")
    print(absent_body)
    
    print("\n✅ Email template test complete!")

def test_send_single_email():
    """Test sending a single email (requires OAuth setup)."""
    print("\n=== Testing Single Email Send ===")
    
    # Get test email address
    test_email = input("Enter email address to send test email to: ").strip()
    if not test_email:
        print("No email provided, skipping test.")
        return
    
    # Create test data
    test_employee = {
        'name': 'Test User',
        'email': test_email,
        'arrival_time': '09:20 AM',
        'minutes_late': 20
    }
    
    date = datetime.datetime.now(pytz.timezone(TIMEZONE)).date()
    start_time = datetime.datetime.strptime("09:00", "%H:%M").time()
    start_dt = datetime.datetime.combine(date, start_time, tzinfo=pytz.timezone(TIMEZONE))
    
    # Send test email
    email_automation = EmailAutomation()
    
    print(f"\nSending test late notification to: {test_email}")
    success = email_automation.send_late_notification(test_employee, date, start_dt)
    
    if success:
        print("✅ Test email sent successfully!")
    else:
        print("❌ Failed to send test email. Check the error messages above.")

def test_batch_notifications():
    """Test batch notification logic (without actually sending)."""
    print("\n=== Testing Batch Notification Logic ===")
    
    # Create test attendance data
    test_data = {
        'IT Dept': {
            'present': [
                {'name': 'John Doe', 'email': 'john@example.com', 'arrival_time': '08:55 AM'}
            ],
            'late': [
                {'name': 'Jane Smith', 'email': 'jane@example.com', 'arrival_time': '09:15 AM', 'minutes_late': 15},
                {'name': 'Bob Wilson', 'email': 'bob@example.com', 'arrival_time': '09:25 AM', 'minutes_late': 25}
            ],
            'absent': [
                {'name': 'Alice Brown', 'email': 'alice@example.com'}
            ]
        },
        'HR Dept': {
            'present': [],
            'late': [
                {'name': 'Charlie Davis', 'email': 'charlie@example.com', 'arrival_time': '10:10 AM', 'minutes_late': 10}
            ],
            'absent': [
                {'name': 'Eve Johnson', 'email': 'eve@example.com'},
                {'name': 'Frank Miller', 'email': 'frank@example.com'}
            ]
        }
    }
    
    # Count notifications
    total_late = sum(len(dept['late']) for dept in test_data.values())
    total_absent = sum(len(dept['absent']) for dept in test_data.values())
    
    print(f"Test data contains:")
    print(f"- {total_late} late employees")
    print(f"- {total_absent} absent employees")
    print(f"Total emails that would be sent: {total_late + total_absent}")
    
    print("\nDepartment breakdown:")
    for dept, data in test_data.items():
        print(f"\n{dept}:")
        print(f"  Late: {len(data['late'])} employees")
        for emp in data['late']:
            print(f"    - {emp['name']} ({emp['email']}) - {emp['minutes_late']} min late")
        print(f"  Absent: {len(data['absent'])} employees")
        for emp in data['absent']:
            print(f"    - {emp['name']} ({emp['email']})")
    
    print("\n✅ Batch notification logic test complete!")

def main():
    """Main test function."""
    print("=== Email Automation Test Suite ===")
    print("This will test the email automation functionality")
    
    while True:
        print("\nSelect a test:")
        print("1. Test email templates (no emails sent)")
        print("2. Send a test email (requires OAuth setup)")
        print("3. Test batch notification logic (no emails sent)")
        print("4. Exit")
        
        choice = input("\nEnter your choice (1-4): ").strip()
        
        if choice == '1':
            test_email_templates()
        elif choice == '2':
            test_send_single_email()
        elif choice == '3':
            test_batch_notifications()
        elif choice == '4':
            print("Exiting tests...")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main() 