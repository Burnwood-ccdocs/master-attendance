#!/usr/bin/env python3
"""
Debug script to examine the data structure and identify issues.
"""

import datetime
import pytz
from generate_report import AttendanceReportGenerator
from config import TIMEZONE

def debug_data_structure():
    """Debug the data structure to understand the index error."""
    print("=== DEBUGGING DATA STRUCTURE ===")
    
    try:
        # Initialize the reporter
        reporter = AttendanceReportGenerator()
        
        # Get today's date
        today = datetime.datetime.now(pytz.timezone(TIMEZONE)).date()
        
        # Fetch WebWork data
        webwork_data = reporter.fetch_webwork_data(today.strftime('%Y-%m-%d'))
        if not webwork_data:
            print("No WebWork data available")
            return
        
        first_entries = reporter.get_first_check_in_times(webwork_data, today)
        
        # Test with IT Dept
        test_dept = "IT Dept"
        print(f"\nTesting with department: {test_dept}")
        
        # Get department employees
        department_employees = reporter.get_department_employees_from_webwork(webwork_data)
        employees = department_employees.get(test_dept, [])
        
        print(f"Found {len(employees)} employees for {test_dept}")
        
        if employees:
            # Show first employee structure
            print(f"First employee structure: {employees[0]}")
            
            # Calculate statuses
            start_dt = datetime.datetime.combine(today, datetime.datetime.strptime("09:00", "%H:%M").time(), tzinfo=reporter.tz)
            present, late, initially_absent, absent = reporter._calculate_statuses(employees, first_entries, start_dt)
            
            print(f"\nStatus counts:")
            print(f"  Present: {len(present)}")
            print(f"  Late: {len(late)}")
            print(f"  Initially Absent: {len(initially_absent)}")
            print(f"  Absent: {len(absent)}")
            
            if present:
                print(f"\nFirst present employee structure: {present[0]}")
            if late:
                print(f"First late employee structure: {late[0]}")
            if absent:
                print(f"First absent employee structure: {absent[0]}")
            
            # Test table building
            print(f"\nTesting table building...")
            from generate_report import SlackNotifier
            notifier = SlackNotifier()
            
            # Test present table
            if present:
                print("Testing present table:")
                result = notifier._build_ascii_table(["Name", "Email", "Time"], present)
                print(result)
            
            # Test late table
            if late:
                print("Testing late table:")
                result = notifier._build_ascii_table(["Name", "Email", "Time", "Min Late"], late)
                print(result)
            
            # Test absent table
            if absent:
                print("Testing absent table:")
                result = notifier._build_ascii_table(["Name", "Email"], absent)
                print(result)
        
    except Exception as e:
        print(f"Error during debugging: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_data_structure() 