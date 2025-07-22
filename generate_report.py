import datetime
import pytz
import gspread
from google.oauth2.service_account import Credentials
import requests
import base64
import os
import sys
from config import (
    WEBWORK_API_URL,
    WEBWORK_USERS_API_URL,
    WEBWORK_API_KEY,
    WEBWORK_API_USER,
    GOOGLE_SERVICE_ACCOUNT_FILE,
    GOOGLE_SHEET_ID,
    DEPARTMENTS_CONFIG,
    TIMEZONE,
    DEPARTMENT_START_TIMES,
    DEFAULT_START_TIME,
    SLACK_BOT_TOKEN,
    SLACK_CHANNEL_ID,
    EMAIL_AUTOMATION_ENABLED
)
import time
import gspread.utils
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from email_automation import EmailAutomation
import json

class SlackNotifier:
    """Handles formatting and sending Slack messages."""
    def __init__(self):
        if not SLACK_BOT_TOKEN or not SLACK_CHANNEL_ID:
            print("Warning: Slack credentials not found. Slack notifications will be disabled.")
            self.client = None
        else:
            self.client = WebClient(token=SLACK_BOT_TOKEN)

    def _build_ascii_table(self, headers, rows):
        """Builds a formatted ASCII table inside a code block."""
        if not rows:
            return "None."
        
        try:
            # Convert dictionary rows to list format for table building
            formatted_rows = []
            for row in rows:
                if isinstance(row, dict):
                    # Handle dictionary format from _calculate_statuses
                    if 'name' in row and 'email' in row:
                        if 'arrival_time' in row:
                            if 'minutes_late' in row:
                                # Late/initially_absent format: name, email, time, minutes_late
                                formatted_rows.append([row['name'], row['email'], row['arrival_time'], str(row['minutes_late'])])
                            else:
                                # Present format: name, email, time
                                formatted_rows.append([row['name'], row['email'], row['arrival_time']])
                        else:
                            # Absent format: name, email
                            formatted_rows.append([row['name'], row['email']])
                    else:
                        # Fallback for other dictionary formats
                        formatted_rows.append([str(v) for v in row.values()])
                else:
                    # Already in list format
                    formatted_rows.append(row)
            
            if not formatted_rows:
                return "None."
            
            # Ensure all rows have the same number of columns as headers
            max_cols = len(headers)
            normalized_rows = []
            for row in formatted_rows:
                # Pad or truncate row to match header length
                normalized_row = []
                for i in range(max_cols):
                    if i < len(row):
                        normalized_row.append(str(row[i]))
                    else:
                        normalized_row.append("")  # Empty cell for missing data
                normalized_rows.append(normalized_row)
            
            # Calculate column widths
            col_widths = [len(h) for h in headers]
            for row in normalized_rows:
                for i, cell in enumerate(row):
                    if i < len(col_widths):
                        col_widths[i] = max(col_widths[i], len(str(cell)))

            # Build table
            header_line = " | ".join(h.ljust(col_widths[i]) for i, h in enumerate(headers))
            separator_line = "-|-".join("-" * col_widths[i] for i in range(len(headers)))
            body_lines = [" | ".join(str(cell).ljust(col_widths[i]) for i, cell in enumerate(row)) for row in normalized_rows]
            
            return "```\n" + "\n".join([header_line, separator_line] + body_lines) + "\n```"
        
        except Exception as e:
            print(f"  Error building table: {e}")
            # Return a simple fallback
            return f"Error building table: {str(e)}"

    def send_consolidated_report(self, report_title, aggregated_data, run_type):
        """Builds and sends a single Slack message for multiple departments."""
        if not self.client or not aggregated_data:
            print("  No data to send to Slack.")
            return

        message_parts = [report_title]

        for dept, data in aggregated_data.items():
            try:
                message_parts.append(f"\n\n--- *{dept}* ---")
                
                if run_type == 'morning':
                    message_parts.append("\n*On-time Arrivals*")
                    message_parts.append(self._build_ascii_table(["Name", "Email", "Time"], data['present']))
                    message_parts.append("\n*Late Arrivals (5-30 min)*")
                    message_parts.append(self._build_ascii_table(["Name", "Email", "Time", "Min Late"], data['late']))
                elif run_type == 'eod':
                    message_parts.append("\n*On-time Arrivals*")
                    message_parts.append(self._build_ascii_table(["Name", "Email", "Time"], data['present']))
                    message_parts.append("\n*Late Arrivals (5-30 min)*")
                    message_parts.append(self._build_ascii_table(["Name", "Email", "Time", "Min Late"], data['late']))
                    message_parts.append("\n*Absent*")
                    message_parts.append(self._build_ascii_table(["Name", "Email"], data['absent']))
            except Exception as e:
                print(f"  Error building table for {dept}: {e}")
                # Add a simple fallback message
                message_parts.append(f"\n*Error building report for {dept}: {str(e)}*")

        final_message = "\n".join(message_parts)
        
        try:
            self.client.chat_postMessage(channel=SLACK_CHANNEL_ID, text=final_message)
            print(f"  Successfully sent consolidated Slack report.")
        except SlackApiError as e:
            print(f"  Error sending consolidated Slack report: {e.response['error']}", file=sys.stderr)


class AttendanceReportGenerator:
    def __init__(self):
        self.tz = pytz.timezone(TIMEZONE)
        self.gc = self._authenticate_google_sheets()
        self.spreadsheet = self.gc.open_by_key(GOOGLE_SHEET_ID)
        self.user_cache = {}
        self.attendance_options = ["Present", "Late", "Absent"]
        
        # Color definitions for formatting
        self.header_color = {"red": 0.26, "green": 0.44, "blue": 0.76}  # Blue
        self.present_color = {"red": 0.77, "green": 0.93, "blue": 0.80}   # Green
        self.late_color = {"red": 1.0, "green": 0.94, "blue": 0.60}     # Yellow
        self.absent_color = {"red": 1.0, "green": 0.77, "blue": 0.80}     # Red
        self.slack_notifier = SlackNotifier()
        
        # Initialize email automation
        self.email_automation = EmailAutomation() if EMAIL_AUTOMATION_ENABLED else None
        
        # Populate user cache on initialization
        self.fetch_user_info()

    def _authenticate_google_sheets(self):
        """Authenticate with Google Sheets using service account."""
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        try:
            # Handle UTF-8 BOM in the service account file
            with open(GOOGLE_SERVICE_ACCOUNT_FILE, 'r', encoding='utf-8-sig') as f:
                service_account_info = json.load(f)
            creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
            return gspread.authorize(creds)
        except Exception as e:
            print(f"Error authenticating with Google Sheets: {e}")
            return None

    def get_auth_header(self):
        """Get Basic Authentication header for WebWork API."""
        credentials = f"{WEBWORK_API_USER}:{WEBWORK_API_KEY}"
        encoded_credentials = base64.b64encode(credentials.encode()).decode()
        return {"Authorization": f"Basic {encoded_credentials}"}

    def fetch_webwork_data(self, date):
        """Fetch attendance data from WebWork API."""
        try:
            response = requests.get(
                WEBWORK_API_URL,
                params={"start_date": date, "end_date": date},
                headers=self.get_auth_header(),
            )
            response.raise_for_status()
            # Handle UTF-8 BOM if present
            response.encoding = 'utf-8-sig'
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching WebWork data: {e}", file=sys.stderr)
            return None

    def fetch_user_info(self):
        """Fetch user information from WebWork API and populate cache."""
        try:
            response = requests.get(WEBWORK_USERS_API_URL, headers=self.get_auth_header())
            response.raise_for_status()
            # Handle UTF-8 BOM if present
            response.encoding = 'utf-8-sig'
            users = response.json()
            for user in users:
                if 'email' in user:
                    self.user_cache[user['email']] = user.get('fullname', user['email'])
            print(f"Successfully cached info for {len(self.user_cache)} users.")
        except requests.exceptions.RequestException as e:
            print(f"Error fetching user info: {e}", file=sys.stderr)

    def get_user_name(self, email):
        """Get user's full name from cache, falling back to email."""
        return self.user_cache.get(email, email)

    def get_department_employees_from_webwork(self, webwork_data):
        """Build department rosters from WebWork data based on project mappings."""
        department_employees = {dept: [] for dept in DEPARTMENTS_CONFIG.keys()}
        
        if not webwork_data or "dateReport" not in webwork_data:
            return department_employees

        for report in webwork_data["dateReport"]:
            email = report.get("email")
            name = self.get_user_name(email)
            
            if not email:
                continue

            for project in report.get("projects", []):
                project_name = project.get("projectName")
                for dept, proj_list in DEPARTMENTS_CONFIG.items():
                    if project_name in proj_list:
                        if not any(e['email'] == email for e in department_employees[dept]):
                            department_employees[dept].append({"name": name, "email": email})
        
        return {dept: emps for dept, emps in department_employees.items() if emps}

    def get_first_check_in_times(self, webwork_data, target_date):
        """Get the first check-in time for each user."""
        first_entries = {}
        if not webwork_data or "dateReport" not in webwork_data:
            return first_entries

        for report in webwork_data["dateReport"]:
            email = report.get("email")
            if not email:
                continue
            
            # Normalize email for consistent matching
            normalized_email = email.lower().strip()
            
            first_entry_time = None
            for project in report.get("projects", []):
                for task in project.get("tasks", []):
                    for entry in task.get("timeEntries", []):
                        try:
                            begin_time = datetime.datetime.strptime(entry["beginDatetime"], "%H:%M").time()
                            entry_datetime = datetime.datetime.combine(target_date, begin_time, tzinfo=self.tz)
                            if first_entry_time is None or entry_datetime < first_entry_time:
                                first_entry_time = entry_datetime
                        except (ValueError, TypeError):
                            continue
            if first_entry_time:
                first_entries[normalized_email] = first_entry_time
        return first_entries

    def _calculate_statuses(self, employees, first_entries, start_dt):
        """Calculates detailed statuses for a list of employees."""
        present, late, initially_absent_late, absent = [], [], [], []

        for emp in employees:
            normalized_email = emp['email'].lower().strip()
            if normalized_email in first_entries:
                check_in = first_entries[normalized_email]
                minutes_late = int((check_in - start_dt).total_seconds() / 60)
                
                emp_details = {
                    'name': emp['name'], 
                    'email': emp['email'],
                    'arrival_time': check_in.strftime("%I:%M %p"),
                    'minutes_late': minutes_late
                }
                
                if minutes_late <= 5:
                    present.append(emp_details)
                elif minutes_late <= 30:
                    late.append(emp_details)
                else:
                    # Anyone more than 30 minutes late is considered absent
                    absent.append({
                        'name': emp['name'],
                        'email': emp['email']
                    })
            else:
                # Ensure absent employees have the same structure as others
                absent.append({
                    'name': emp['name'],
                    'email': emp['email']
                })
        
        return present, late, initially_absent_late, absent

    def run_consolidated_report(self, departments_to_process, date, run_type):
        """
        Processes a list of departments, updates their sheets, and sends one
        aggregated Slack report.
        """
        print(f"--- Starting {run_type.upper()} consolidated run for {len(departments_to_process)} departments ---")
        
        webwork_data = self.fetch_webwork_data(date.strftime('%Y-%m-%d'))
        if not webwork_data:
            print("Aborting run due to WebWork API failure.")
            # Send a notification about the API failure
            error_title = f":warning: WebWork API Error - {run_type.title()} Report - {date.strftime('%Y-%m-%d')}"
            error_message = "Unable to fetch attendance data from WebWork API. Please check the API status and credentials."
            try:
                if self.slack_notifier.client:
                    self.slack_notifier.client.chat_postMessage(channel=SLACK_CHANNEL_ID, text=error_message)
                    print("  Sent API error notification to Slack.")
            except Exception as e:
                print(f"  Could not send error notification: {e}")
            return
        
        first_entries = self.get_first_check_in_times(webwork_data, date)
        aggregated_results = {}

        for department in departments_to_process:
            print(f"  Processing {department}...")
            status_data = self._update_sheet_and_get_statuses(department, date, webwork_data, first_entries)
            if status_data:
                aggregated_results[department] = status_data
                print(f"    Found {len(status_data['present'])} present, {len(status_data['late'])} late, {len(status_data['absent'])} absent")
            else:
                print(f"    No data for {department}")

        if not aggregated_results:
            print("  No data found for any departments.")
            return

        # Send one Slack message with all the results
        report_title = f":loudspeaker: {run_type.title()} Attendance Report - {date.strftime('%Y-%m-%d')}"
        self.slack_notifier.send_consolidated_report(report_title, aggregated_results, run_type)
        
        # Send email notifications to late and absent employees
        if self.email_automation and run_type == 'eod':  # Send emails only during EOD report
            print("\n--- Email Notifications ---")
            try:
                # Prepare department start times
                dept_start_times = {}
                for dept in aggregated_results.keys():
                    start_time_str = DEPARTMENT_START_TIMES.get(dept, DEFAULT_START_TIME)
                    start_dt = datetime.datetime.combine(date, datetime.datetime.strptime(start_time_str, "%H:%M").time(), tzinfo=self.tz)
                    dept_start_times[dept] = start_dt
                
                # Send batch notifications
                self.email_automation.send_batch_notifications(aggregated_results, date, dept_start_times)
            except Exception as e:
                print(f"Error sending email notifications: {e}")
        
        print(f"--- Consolidated {run_type.upper()} run complete ---")

    def _update_sheet_and_get_statuses(self, department, date, webwork_data, first_entries):
        """Helper to contain the logic for processing one department's sheet."""
        try:
            # Get department employees from WebWork data
            department_employees = self.get_department_employees_from_webwork(webwork_data)
            employees = department_employees.get(department, [])
            
            if not employees:
                print(f"  No employees found for {department}")
                return None
            
            # Calculate start time for this department
            start_dt = datetime.datetime.combine(date, datetime.datetime.strptime(DEPARTMENT_START_TIMES.get(department, DEFAULT_START_TIME), "%H:%M").time(), tzinfo=self.tz)
            
            # Calculate statuses
            present, late, initially_absent, absent = self._calculate_statuses(employees, first_entries, start_dt)
            
            # Update the Google Sheet
            self._update_department_sheet(department, date, employees, first_entries, start_dt)
            
            return {
                "present": present, 
                "late": late, 
                "initially_absent": initially_absent, 
                "absent": absent
            }
        except Exception as e:
            print(f"  Error processing {department}: {e}", file=sys.stderr)
            return None

    def _update_department_sheet(self, department, date, employees, first_entries, start_dt):
        """Update the Google Sheet for a specific department with attendance data."""
        try:
            # Get or create the worksheet for this department
            try:
                worksheet = self.spreadsheet.worksheet(department)
                print(f"    Using existing worksheet for {department}")
            except:
                # Create new worksheet if it doesn't exist
                worksheet = self.spreadsheet.add_worksheet(title=department, rows=1000, cols=10)
                print(f"    Created new worksheet for {department}")
            
            # Get existing data to preserve structure
            existing_data = worksheet.get_all_values()
            date_str = date.strftime('%Y-%m-%d')
            
            if not existing_data:
                # If sheet is empty, create initial structure
                header_row = ["Name", "Email"]
                sheet_data = [header_row]
                
                # Add employee data with just name and email
                for emp in employees:
                    row = [emp['name'], emp['email']]
                    sheet_data.append(row)
                
                worksheet.update(sheet_data)
                print(f"    Created initial structure for {department} with {len(employees)} employees")
                return
            
            # Work with existing data structure
            print(f"    Working with existing data structure for {department}")
            
            # Find the date column or create new one
            header_row = existing_data[0] if existing_data else ["Name", "Email"]
            
            # Check if date column already exists
            date_col_index = None
            for i, col in enumerate(header_row):
                if col == date_str:
                    date_col_index = i
                    break
            
            if date_col_index is None:
                # Add new date column
                date_col_index = len(header_row)
                # Extend header row
                header_row.append(date_str)
                # Extend all data rows with empty cells
                for i in range(1, len(existing_data)):
                    while len(existing_data[i]) < len(header_row):
                        existing_data[i].append("")
            
            # Update header row
            existing_data[0] = header_row
            
            # Create a mapping of email to row index for existing data
            email_to_row = {}
            for i, row in enumerate(existing_data[1:], 1):
                if len(row) > 1:
                    email = row[1] if len(row) > 1 else ""
                    if email:  # Only add non-empty emails
                        email_to_row[email.lower().strip()] = i  # Normalize email for matching
            
            print(f"    Found {len(email_to_row)} existing employees in sheet")
            
            # Create a mapping of WebWork employees for quick lookup
            webwork_employees = {emp['email'].lower().strip(): emp for emp in employees}
            print(f"    Found {len(webwork_employees)} employees with WebWork activity")
            
            # Update existing rows and add new employees
            updated_employees = set()
            matched_employees = 0
            new_employees = 0
            preserved_employees = 0
            
            # Process all existing employees in the sheet
            for i, row in enumerate(existing_data[1:], 1):
                if len(row) > 1:
                    email = row[1] if len(row) > 1 else ""
                    if email:
                        normalized_email = email.lower().strip()
                        
                        # Ensure row has enough columns
                        while len(existing_data[i]) < len(header_row):
                            existing_data[i].append("")
                        
                        if normalized_email in webwork_employees:
                            # Update with WebWork data
                            emp = webwork_employees[normalized_email]
                            existing_data[i][0] = emp['name']  # Update name
                            
                            # Update attendance status
                            if normalized_email in first_entries:
                                check_in = first_entries[normalized_email]
                                minutes_late = int((check_in - start_dt).total_seconds() / 60)
                                
                                if minutes_late <= 5:
                                    status = "Present"
                                elif minutes_late <= 30:
                                    status = "Late"
                                else:
                                    status = "Absent"
                            else:
                                status = "Absent"
                            
                            existing_data[i][date_col_index] = status
                            updated_employees.add(normalized_email)
                            matched_employees += 1
                            print(f"      ✓ Updated {emp['name']} ({email}) - {status}")
                        else:
                            # Employee exists in sheet but no WebWork activity today
                            # Mark as absent for today
                            existing_data[i][date_col_index] = "Absent"
                            preserved_employees += 1
                            print(f"      - Preserved {row[0]} ({email}) - Absent (no activity)")
            
            # Add new employees from WebWork that don't exist in sheet
            for emp in employees:
                normalized_email = emp['email'].lower().strip()
                if normalized_email not in email_to_row:
                    new_employees += 1
                    new_row = [emp['name'], emp['email']]
                    # Add empty cells for existing date columns
                    for i in range(2, len(header_row)):
                        if i == date_col_index:
                            # Add attendance status for current date
                            if normalized_email in first_entries:
                                check_in = first_entries[normalized_email]
                                minutes_late = int((check_in - start_dt).total_seconds() / 60)
                                
                                if minutes_late <= 5:
                                    status = "Present"
                                elif minutes_late <= 30:
                                    status = "Late"
                                else:
                                    status = "Absent"
                            else:
                                status = "Absent"
                            new_row.append(status)
                        else:
                            new_row.append("")
                    
                    existing_data.append(new_row)
                    print(f"      + Added {emp['name']} ({emp['email']}) - {status}")
            
            # Update the sheet with all data
            worksheet.clear()
            worksheet.update(existing_data)
            
            # Apply formatting
            self._format_header(worksheet)
            self._apply_conditional_formatting(worksheet)
            
            # Add dropdown validation to all date columns (except Name and Email)
            for col_num in range(3, len(header_row) + 1):  # Start from column C (3)
                self._add_dropdown_validation(worksheet, col_num)
            
            print(f"    Updated sheet for {department}: {matched_employees} updated, {preserved_employees} preserved, {new_employees} new, {len(existing_data)-1} total employees")
            
        except Exception as e:
            print(f"    Error updating sheet for {department}: {e}", file=sys.stderr)
            raise

    def process_department_and_notify(self, department, date, run_type, webwork_data, first_entries):
        """Process a single department and send individual notification."""
        print(f"Processing {department} for {run_type} report...")
        
        start_dt = datetime.datetime.combine(date, datetime.datetime.strptime(DEPARTMENT_START_TIMES.get(department, DEFAULT_START_TIME), "%H:%M").time(), tzinfo=self.tz)
        
        # Get department employees from WebWork data
        department_employees = self.get_department_employees_from_webwork(webwork_data)
        employees = department_employees.get(department, [])
        
        if not employees:
            print(f"  No employees found for {department}")
            return
        
        # Calculate statuses
        present, late, initially_absent, absent = self._calculate_statuses(employees, first_entries, start_dt)
        
        # Update the Google Sheet
        self._update_department_sheet(department, date, employees, first_entries, start_dt)
        
        # Create status data for Slack notification
        status_data = {
            "present": present,
            "late": late, 
            "initially_absent": initially_absent,
            "absent": absent
        }
        
        # Send individual department notification
        report_title = f":loudspeaker: {run_type.title()} Attendance Report - {department} - {date.strftime('%Y-%m-%d')}"
        self.slack_notifier.send_consolidated_report(report_title, {department: status_data}, run_type)

    def generate_report(self, date, department):
        """Wrapper to fetch data and process a single department for the initial run."""
        webwork_data = self.fetch_webwork_data(date.strftime('%Y-%m-%d'))
        first_entries = self.get_first_check_in_times(webwork_data, date)
        self.process_department_and_notify(department, date, 'morning', webwork_data, first_entries)

    def run_end_of_day_update(self, date, department):
        """Wrapper to fetch data and process a single department for the EOD run."""
        webwork_data = self.fetch_webwork_data(date.strftime('%Y-%m-%d'))
        first_entries = self.get_first_check_in_times(webwork_data, date)
        self.process_department_and_notify(department, date, 'eod', webwork_data, first_entries)

    def _format_header(self, worksheet):
        """Applies bold text and background color to the header row."""
        try:
            # Get current data to determine column range
            data = worksheet.get_all_values()
            if not data:
                print("  No data to format header for.")
                return
            
            header_row = data[0]
            col_count = len(header_row)
            
            # Format the entire header row
            range_notation = f'A1:{self._col_to_a1(col_count)}1'
            worksheet.format(range_notation, {
                "backgroundColor": self.header_color,
                "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}, "bold": True}
            })
            print(f"  Formatted header row ({col_count} columns).")
        except Exception as e:
            print(f"  Could not format header. Error: {e}", file=sys.stderr)

    def _apply_conditional_formatting(self, worksheet):
        """Applies color-coded conditional formatting for attendance statuses."""
        try:
            # Get the current data to determine column range
            data = worksheet.get_all_values()
            if not data or len(data) < 2:
                print("  No data to apply conditional formatting to.")
                return
            
            header_row = data[0]
            # Apply formatting to all date columns (columns 3 and beyond)
            date_columns = []
            for i, col in enumerate(header_row):
                if i >= 2:  # Skip Name and Email columns
                    date_columns.append(i)
            
            if not date_columns:
                print("  No date columns found for conditional formatting.")
                return
            
            rules = []
            rule_index = 0
            
            for col_index in date_columns:
                # Present - Green
                rules.append({
                    "ranges": [{"sheetId": worksheet.id, "startRowIndex": 1, "startColumnIndex": col_index, "endColumnIndex": col_index + 1}],
                "booleanRule": {
                    "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Present"}]},
                    "format": {"backgroundColor": self.present_color}
                }
                })
                
                # Late - Yellow
                rules.append({
                    "ranges": [{"sheetId": worksheet.id, "startRowIndex": 1, "startColumnIndex": col_index, "endColumnIndex": col_index + 1}],
                "booleanRule": {
                    "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Late"}]},
                    "format": {"backgroundColor": self.late_color}
                }
                })
                
                # Absent - Red
                rules.append({
                    "ranges": [{"sheetId": worksheet.id, "startRowIndex": 1, "startColumnIndex": col_index, "endColumnIndex": col_index + 1}],
                "booleanRule": {
                    "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "Absent"}]},
                    "format": {"backgroundColor": self.absent_color}
                }
                })
        
            if rules:
                requests = [{'addConditionalFormatRule': {'rule': rule, 'index': i}} for i, rule in enumerate(rules)]
                worksheet.spreadsheet.batch_update({'requests': requests})
                print(f"  Applied conditional formatting to {len(date_columns)} date columns.")
            else:
                print("  No conditional formatting rules to apply.")
                
        except Exception as e:
            print(f"  Could not apply conditional formatting. Error: {e}", file=sys.stderr)

    def _col_to_a1(self, col):
        """Converts a 1-based column number to A1 notation (e.g., 1 -> A, 27 -> AA)."""
        if col < 1:
            raise ValueError("Column number must be 1 or greater.")
        
        result = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _add_dropdown_validation(self, worksheet, col_num):
        """Add dropdown data validation to a column using a compatible method."""
        try:
            validation_rule = {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": option} for option in self.attendance_options]
                },
                "showCustomUi": True,
                "strict": True
            }
            
            col_letter = self._col_to_a1(col_num)
            requests = [{
                "setDataValidation": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": 1,
                        "startColumnIndex": col_num - 1,
                        "endColumnIndex": col_num,
                    },
                    "rule": validation_rule,
                }
            }]
            worksheet.spreadsheet.batch_update({"requests": requests})
            print(f"  Added dropdown validation to column {col_letter}.")
        except Exception as e:
            print(f"  Could not set dropdown validation. Error: {e}", file=sys.stderr)

def main():
    """Main function to trigger the report."""
    parser = argparse.ArgumentParser(description="Generate and update attendance reports.")
    parser.add_argument(
        '--test-week',
        action='store_true',
        help="Run the report for every day of the current week (Mon-Today)."
    )
    args = parser.parse_args()

    reporter = AttendanceReportGenerator()
    
    if args.test_week:
        print("--- Running Weekly Test Mode (with Sheet Updates & Slack Notifications) ---")
        today = datetime.datetime.now(pytz.timezone(TIMEZONE)).date()
        start_of_week = today - datetime.timedelta(days=today.weekday())
        
        current_day = start_of_week
        while current_day <= today:
            print(f"\n>>> Processing FULL DAY report for: {current_day.strftime('%A, %Y-%m-%d')}")
            # In test mode, we process all departments for the day.
            for dept in DEPARTMENTS_CONFIG.keys():
                reporter.generate_report(current_day, dept)
                reporter.run_end_of_day_update(current_day, dept)
            current_day += datetime.timedelta(days=1)
        print("\n--- Weekly Test Mode Complete ---")
    else:
        print("This script is now designed to be run via the main.py scheduler or with the --test-week flag.")

if __name__ == "__main__":
    import argparse
    main() 