import requests
import datetime
import pytz
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from config import *
import base64
import sys
from apscheduler.schedulers.blocking import BlockingScheduler

class AttendanceTracker:
    def __init__(self):
        self.webwork_api_url = WEBWORK_API_URL
        self.webwork_users_api_url = WEBWORK_USERS_API_URL
        self.webwork_api_user = WEBWORK_API_USER
        self.webwork_api_key = WEBWORK_API_KEY
        self.slack_client = WebClient(token=SLACK_BOT_TOKEN)
        self.timezone = pytz.timezone(TIMEZONE)
        self.hr_project_name = HR_PROJECT_NAME
        self.user_cache = {}  # Cache for user information

    def get_today_date(self):
        """Get today's date in YYYY-MM-DD format"""
        return datetime.datetime.now(self.timezone).strftime("%Y-%m-%d")

    def parse_time(self, time_str, date):
        """Parse time string in HH:MM format and combine with date"""
        try:
            hour, minute = map(int, time_str.split(':'))
            time = datetime.time(hour, minute)
            return datetime.datetime.combine(date, time, tzinfo=self.timezone)
        except (ValueError, TypeError) as e:
            print(f"Error parsing time {time_str}: {e}")
            return None

    def get_auth_header(self):
        """Get Basic Authentication header"""
        credentials = f"{self.webwork_api_user}:{self.webwork_api_key}"
        encoded_credentials = base64.b64encode(credentials.encode()).decode()
        return {"Authorization": f"Basic {encoded_credentials}"}

    def fetch_webwork_data(self, date):
        """Fetch attendance data from WebWork API using Basic Authentication"""
        try:
            response = requests.get(
                self.webwork_api_url,
                params={
                    "start_date": date,
                    "end_date": date
                },
                headers=self.get_auth_header()
            )
            response.raise_for_status()
            # Handle UTF-8 BOM if present
            response.encoding = 'utf-8-sig'
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching WebWork data: {e}")
            if hasattr(e.response, 'text'):
                print(f"Response content: {e.response.text}")
            return None

    def fetch_user_info(self):
        """Fetch user information from WebWork API"""
        try:
            response = requests.get(
                self.webwork_users_api_url,
                headers=self.get_auth_header()
            )
            response.raise_for_status()
            # Handle UTF-8 BOM if present
            response.encoding = 'utf-8-sig'
            users = response.json()
            for user in users:
                if 'email' in user:
                    self.user_cache[user['email']] = {
                        'name': user.get('fullname', user['email']),
                        'email': user['email']
                    }
            print(f"Successfully fetched information for {len(self.user_cache)} users")
        except requests.exceptions.RequestException as e:
            print(f"Error fetching user information: {e}")
            if hasattr(e.response, 'text'):
                print(f"Response content: {e.response.text}")

    def get_user_name(self, email):
        """Get user's full name from cache"""
        if email in self.user_cache and self.user_cache[email]['name']:
            return self.user_cache[email]['name']
        return email  # Fallback to email if name not found

    def get_hr_team_members(self, data):
        """Extract HR team members from the Internal CCDOCS-HR project"""
        hr_team_emails = set()
        
        if not data or "dateReport" not in data:
            print("No data received from WebWork API")
            return hr_team_emails

        for report in data["dateReport"]:
            for project in report.get("projects", []):
                if project.get("projectName") == self.hr_project_name:
                    email = report.get("email")
                    if email:
                        hr_team_emails.add(email)
                    break

        # Include manually specified HR members
        try:
            from config import ADDITIONAL_HR_EMAILS
            hr_team_emails.update(ADDITIONAL_HR_EMAILS)
        except ImportError:
            # Constant not defined; ignore
            pass

        # Exclude specific emails
        try:
            from config import EXCLUDED_EMAILS
            hr_team_emails -= set(EXCLUDED_EMAILS)
        except ImportError:
            pass  # No exclusions defined

        if not hr_team_emails:
            print(f"No team members found in project: {self.hr_project_name}")
        else:
            print(f"Found {len(hr_team_emails)} total team members (including additional list, minus exclusions)")

        return list(hr_team_emails)

    def get_joining_times(self, data, *, target_date: datetime.date | None = None):
        """Get joining times for all HR team members (or Absent) for a specific date.

        Parameters
        ----------
        data : dict
            JSON payload from WebWork daily-timeline API.
        target_date : datetime.date, optional
            The date that *data* corresponds to. Defaults to today (in the configured timezone).
        """
        joining_times = []
        hr_team_emails = set(self.get_hr_team_members(data))

        if target_date is None:
            target_date = datetime.datetime.now(self.timezone).date()
        present_emails = set()
        first_entries = {}

        for report in data.get("dateReport", []):
            email = report.get("email")
            if email not in hr_team_emails:
                continue
            first_entry = None
            for project in report.get("projects", []):
                for task in project.get("tasks", []):
                    for time_entry in task.get("timeEntries", []):
                        begin_time = self.parse_time(time_entry["beginDatetime"], target_date)
                        if begin_time and (first_entry is None or begin_time < first_entry):
                            first_entry = begin_time
            if first_entry:
                present_emails.add(email)
                first_entries[email] = first_entry

        for email in hr_team_emails:
            name = self.get_user_name(email)
            if email in first_entries:
                joining_times.append({
                    "email": email,
                    "name": name,
                    "arrival_time": first_entries[email].strftime("%I:%M %p")
                })
            else:
                joining_times.append({
                    "email": email,
                    "name": name,
                    "arrival_time": "Absent"
                })
        return joining_times, first_entries, hr_team_emails

    def get_late_arrivals(self, first_entries):
        """Get late arrivals from first_entries dict"""
        late_arrivals = []
        today_date = datetime.datetime.now(self.timezone).date()
        start_time = datetime.datetime.strptime(START_TIME, "%H:%M").time()
        start_datetime = datetime.datetime.combine(today_date, start_time, tzinfo=self.timezone)
        late_threshold = start_datetime + datetime.timedelta(minutes=LATE_THRESHOLD_MINUTES)
        for email, first_entry in first_entries.items():
            if first_entry > late_threshold:
                minutes_late = int((first_entry - start_datetime).total_seconds() / 60)
                late_arrivals.append({
                    "email": email,
                    "name": self.get_user_name(email),
                    "arrival_time": first_entry.strftime("%I:%M %p"),
                    "minutes_late": minutes_late
                })
        return late_arrivals

    def get_absentees(self, hr_team_emails, first_entries):
        """Get absentees from HR team"""
        absentees = []
        for email in hr_team_emails:
            if email not in first_entries:
                absentees.append({
                    "email": email,
                    "name": self.get_user_name(email)
                })
        return absentees

    def categorize_attendance(self, first_entries, hr_team_emails):
        """Categorise team members into On-time, Late, Very-late (initially absent) and Absent."""
        today_date = datetime.datetime.now(self.timezone).date()

        # Define default cut-off times based on START_TIME
        start_time_default = datetime.datetime.strptime(START_TIME, "%H:%M").time()
        start_datetime_default = datetime.datetime.combine(today_date, start_time_default, tzinfo=self.timezone)

        five_minutes_after_default = start_datetime_default + datetime.timedelta(minutes=5)
        thirty_minutes_after_default = start_datetime_default + datetime.timedelta(minutes=30)
        five_pm_datetime = datetime.datetime.combine(today_date, datetime.time(17, 0), tzinfo=self.timezone)

        # Import per-user overrides lazily to avoid circular import at top
        try:
            from config import CUSTOM_START_TIMES  # type: ignore
        except ImportError:
            CUSTOM_START_TIMES = {}

        on_time = []                    # Logged in <= 5 minutes after start
        late = []                       # Logged in between 5–30 minutes after start
        very_late = []                  # Logged in > 30 minutes after start but before 5 PM
        absentees = []                  # No login or ≥ 5 PM

        for email in hr_team_emails:
            name = self.get_user_name(email)
            first_entry = first_entries.get(email)

            # Determine per-user start datetime (allows late shift workers)
            if email in CUSTOM_START_TIMES:
                custom_start_time = datetime.datetime.strptime(CUSTOM_START_TIMES[email], "%H:%M").time()
                start_dt = datetime.datetime.combine(today_date, custom_start_time, tzinfo=self.timezone)
                start_cutoff = start_dt
                sixty_after = start_dt + datetime.timedelta(minutes=60)
            else:
                start_dt = start_datetime_default
                start_cutoff = start_dt
                sixty_after = start_dt + datetime.timedelta(minutes=60)

            if not first_entry:
                # Never logged in – Absent
                absentees.append({"email": email, "name": name})
                continue

            if first_entry <= start_cutoff:
                on_time.append({
                    "email": email,
                    "name": name,
                    "arrival_time": first_entry.strftime("%I:%M %p")
                })
            elif first_entry < sixty_after:
                minutes_late = int((first_entry - start_dt).total_seconds() / 60)
                late.append({
                    "email": email,
                    "name": name,
                    "arrival_time": first_entry.strftime("%I:%M %p"),
                    "minutes_late": minutes_late
                })
            elif first_entry < five_pm_datetime:
                minutes_late = int((first_entry - start_dt).total_seconds() / 60)
                very_late.append({
                    "email": email,
                    "name": name,
                    "arrival_time": first_entry.strftime("%I:%M %p"),
                    "minutes_late": minutes_late
                })
            else:
                # Logged in at/after 5 PM – still considered Absent
                absentees.append({"email": email, "name": name})

        return on_time, late, very_late, absentees

    def send_slack_report(self, on_time, late, very_late, absentees, *, include_very_late: bool = True):
        """Send attendance report (usually 09:30) including on-time, late and optionally >30-min late, plus absentees."""

        # Helper to build an ASCII table inside a code block
        def build_table(headers, rows):
            col_widths = [len(h) for h in headers]
            for row in rows:
                for i, cell in enumerate(row):
                    col_widths[i] = max(col_widths[i], len(str(cell)))

            # Build separator and header lines
            header_line = " | ".join(h.ljust(col_widths[i]) for i, h in enumerate(headers))
            separator_line = "-|-".join("-" * col_widths[i] for i in range(len(headers)))

            body_lines = [
                " | ".join(str(cell).ljust(col_widths[i]) for i, cell in enumerate(row))
                for row in rows
            ]

            return "```\n" + "\n".join([header_line, separator_line] + body_lines) + "\n```"

        message_parts = [f"<@{SLACK_USER_ID}> *Attendance Report*\n"]

        # On-time section
        message_parts.append("*On-time Arrivals*")
        if on_time:
            ot_rows = [[e["name"], e["email"], e["arrival_time"]] for e in on_time]
            message_parts.append(build_table(["Name", "Email", "Time"], ot_rows))
        else:
            message_parts.append("No on-time arrivals.")

        # Late (5-30 min) section
        message_parts.append("\n*Late Arrivals (5-30 min)*")
        if late:
            late_rows = [[e["name"], e["email"], e["arrival_time"], e["minutes_late"]] for e in late]
            message_parts.append(build_table(["Name", "Email", "Time", "Min Late"], late_rows))
        else:
            message_parts.append("No late arrivals.")

        if include_very_late:
            message_parts.append("\n*Initially Absent but Marked Late (>30 min)*")
            if very_late:
                vl_rows = [[e["name"], e["email"], e["arrival_time"], e["minutes_late"]] for e in very_late]
                message_parts.append(build_table(["Name", "Email", "Time", "Min Late"], vl_rows))
            else:
                message_parts.append("None.")

        # Absentees
        message_parts.append("\n*Absent*")
        if absentees:
            ab_rows = [[e["name"], e["email"]] for e in absentees]
            message_parts.append(build_table(["Name", "Email"], ab_rows))
        else:
            message_parts.append("No absentees.")

        # Combine and send
        final_message = "\n".join(message_parts)
        self.slack_client.chat_postMessage(channel=SLACK_CHANNEL_ID, text=final_message)
        print(f"Successfully sent report to Slack channel {SLACK_CHANNEL_ID}")

    def send_absent_report(self, absentees):
        """Send a dedicated Absent report (used at 10:05 AM)."""

        if not absentees:
            text = f"<@{SLACK_USER_ID}> No absentees as of 10:05 AM!"
        else:
            # Reuse the table builder used in send_slack_report
            def build_table(headers, rows):
                col_widths = [len(h) for h in headers]
                for row in rows:
                    for i, cell in enumerate(row):
                        col_widths[i] = max(col_widths[i], len(str(cell)))

                header_line = " | ".join(h.ljust(col_widths[i]) for i, h in enumerate(headers))
                separator_line = "-|-".join("-" * col_widths[i] for i in range(len(headers)))
                body_lines = [
                    " | ".join(str(cell).ljust(col_widths[i]) for i, cell in enumerate(row))
                    for row in rows
                ]
                return "```\n" + "\n".join([header_line, separator_line] + body_lines) + "\n```"

            table = build_table(["Name", "Email"], [[e["name"], e["email"]] for e in absentees])
            text = f"<@{SLACK_USER_ID}> *Absent Report (10:05 AM)*\n" + table

        self.slack_client.chat_postMessage(channel=SLACK_CHANNEL_ID, text=text)
        print("Sent Absent report to Slack.")

    def send_end_of_day_report(self, very_late, absentees):
        """Send 5 PM report listing those who were initially absent (>30 min) and final absentees."""

        def build_table(headers, rows):
            col_widths = [len(h) for h in headers]
            for row in rows:
                for i, cell in enumerate(row):
                    col_widths[i] = max(col_widths[i], len(str(cell)))

            header_line = " | ".join(h.ljust(col_widths[i]) for i, h in enumerate(headers))
            separator_line = "-|-".join("-" * col_widths[i] for i in range(len(headers)))
            body_lines = [
                " | ".join(str(cell).ljust(col_widths[i]) for i, cell in enumerate(row))
                for row in rows
            ]
            return "```\n" + "\n".join([header_line, separator_line] + body_lines) + "\n```"

        message_parts = [f"<@{SLACK_USER_ID}> *End-of-Day Attendance Summary*\n"]

        # Initially Absent but Marked Late (>30 min)
        message_parts.append("*Marked Late (>30 min)*")
        if very_late:
            vl_rows = [[e["name"], e["email"], e["arrival_time"], e["minutes_late"]] for e in very_late]
            message_parts.append(build_table(["Name", "Email", "Time", "Min Late"], vl_rows))
        else:
            message_parts.append("None.")

        # Absentees
        message_parts.append("\n*Absent*")
        if absentees:
            ab_rows = [[e["name"], e["email"]] for e in absentees]
            message_parts.append(build_table(["Name", "Email"], ab_rows))
        else:
            message_parts.append("No absentees.")

        final_message = "\n".join(message_parts)
        self.slack_client.chat_postMessage(channel=SLACK_CHANNEL_ID, text=final_message)
        print("Sent End-of-Day report to Slack.")

    def run_end_of_day_check(self):
        """Run at 17:00 to send very_late + absentees report."""
        print(f"Starting end-of-day attendance check for {self.get_today_date()}")
        self.fetch_user_info()
        data = self.fetch_webwork_data(self.get_today_date())
        _, first_entries, hr_team_emails = self.get_joining_times(data)
        _on_time, _late, very_late, absentees = self.categorize_attendance(first_entries, hr_team_emails)
        self.send_end_of_day_report(very_late, absentees)
        print("End-of-day attendance check completed")

    def run_daily_check(self):
        """Run the daily attendance check"""
        print(f"Starting daily attendance check for {self.get_today_date()}")
        self.fetch_user_info()
        data = self.fetch_webwork_data(self.get_today_date())
        _, first_entries, hr_team_emails = self.get_joining_times(data)
        on_time, late, very_late, absentees = self.categorize_attendance(first_entries, hr_team_emails)
        self.send_slack_report(on_time, late, very_late, absentees, include_very_late=False)

        # Diagnostic: list projects for vinamrg@ccdocs.com in current report
        vin_projects = self.get_user_projects("vinamrg@ccdocs.com", data)
        if vin_projects:
            print(f"vinamrg@ccdocs.com projects today: {', '.join(vin_projects)}")
        else:
            print("vinamrg@ccdocs.com not found in today's WebWork data.")
        print("Daily attendance check completed")

    def get_user_projects(self, email, data):
        """Return a set of project names the specified user logged time to in the given data report."""
        projects = set()
        if not data or "dateReport" not in data:
            return projects
        for report in data["dateReport"]:
            if report.get("email") == email:
                for project in report.get("projects", []):
                    pname = project.get("projectName")
                    if pname:
                        projects.add(pname)
        return projects

if __name__ == "__main__":
    # Allow running once immediately with --once; otherwise schedule twice daily
    if len(sys.argv) > 1 and sys.argv[1] == "--once":
        tracker = AttendanceTracker()
        tracker.run_daily_check()
    else:
        tracker = AttendanceTracker()

        eastern = pytz.timezone("US/Eastern")

        scheduler = BlockingScheduler(timezone=eastern)

        # Cron-style schedules in Eastern Time
        scheduler.add_job(tracker.run_daily_check, "cron", hour=9, minute=30)
        scheduler.add_job(tracker.run_end_of_day_check, "cron", hour=17, minute=0)
        # Weekly attendance sheet update at 17:30
        def weekly_job():
            from weekly_attendance import WeeklyAttendance  # lazy import to avoid heavy deps at startup
            wa = WeeklyAttendance()
            wa.update_until_date(datetime.datetime.now(eastern).date())

        scheduler.add_job(weekly_job, "cron", hour=17, minute=30)

        # Temporary test: run daily report at 18:44
        def test_report_job():
            tracker.run_daily_check()
            print("Test report sent at 6:44 PM Eastern.")

        print("Attendance tracker scheduler started (US/Eastern):\n - 09:30 full (late) report\n - 17:00 absent + >30-min late report\n - 17:30 weekly sheet update\n Press Ctrl+C to exit.")

        scheduler.start() 