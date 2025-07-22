import datetime
import pytz
import time
import random
from typing import Dict, List

import gspread
from gspread.exceptions import APIError
from gspread_formatting import CellFormat, Color, format_cell_range, batch_updater

from attendance_tracker import AttendanceTracker
from config import (
    TIMEZONE,
    GOOGLE_SERVICE_ACCOUNT_FILE,
    GOOGLE_SHEET_ID,
    START_TIME,
)

# ---------------------------------------------------------------------
# Rate-limit handling
# ---------------------------------------------------------------------
def with_retry(func):
    """Decorator to handle rate-limiting with exponential backoff."""
    def wrapper(*args, **kwargs):
        max_retries = 5
        retry_count = 0
        base_delay = 2  # seconds

        while True:
            try:
                return func(*args, **kwargs)
            except APIError as e:
                error_msg = str(e).lower()
                rate_limited = any(msg in error_msg for msg in [
                    "quota exceeded",
                    "rate limit",
                    "too many requests",
                    "exceeds the limit",
                    "request rate"
                ])

                if not rate_limited or retry_count >= max_retries:
                    raise  # Re-raise if not rate-limited or max retries reached

                retry_count += 1
                # Exponential backoff with jitter
                delay = (base_delay ** retry_count) + random.uniform(0.1, 1.0)
                print(f"⚠️ Rate limit hit. Backing off for {delay:.1f} seconds (attempt {retry_count}/{max_retries})...")
                time.sleep(delay)
    return wrapper

class WeeklyAttendance:
    """Updates the current week's Google Sheet with attendance data.

    If the script is run mid-week it will automatically back-fill previous
    weekdays with real data (instead of marking them Absent)."""

    # Column order in target worksheet
    HEADER = [
        "Employee Name",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
    ]

    # Colours (RGB 0-1 range)
    YELLOW = Color(1.0, 1.0, 0.6)
    RED = Color(1.0, 0.6, 0.6)
    WHITE = Color(1.0, 1.0, 1.0)

    def __init__(self):
        self.tz = pytz.timezone(TIMEZONE)
        self.tracker = AttendanceTracker()

        # Authenticate with Google Sheets
        self.gc = gspread.service_account(GOOGLE_SERVICE_ACCOUNT_FILE)
        self.spread = self.gc.open_by_key(GOOGLE_SHEET_ID)

    # ---------------------------------------------------------------------
    # Sheet helpers
    # ---------------------------------------------------------------------
    def _week_bounds(self, date: datetime.date):
        """Return Monday and Friday for the ISO week containing *date*."""
        monday = date - datetime.timedelta(days=date.weekday())
        friday = monday + datetime.timedelta(days=4)
        return monday, friday

    def _sheet_name_for_week(self, monday: datetime.date, friday: datetime.date) -> str:
        return f"{monday.strftime('%d/%m/%Y')} - {friday.strftime('%d/%m/%Y')}"

    @with_retry
    def _get_or_create_week_sheet(self, monday: datetime.date, friday: datetime.date):
        name = self._sheet_name_for_week(monday, friday)
        try:
            worksheet = self.spread.worksheet(name)
        except gspread.WorksheetNotFound:
            worksheet = self.spread.add_worksheet(title=name, rows="100", cols="10")
            # Write headers
            worksheet.update([self.HEADER])
        return worksheet

    @with_retry
    def _ensure_employee_row(self, worksheet, employee_name: str) -> int:
        """Ensure a row exists for *employee_name* and return its 1-indexed row number."""
        names = worksheet.col_values(1)
        if employee_name in names:
            return names.index(employee_name) + 1  # 1-indexed
        # Append new employee row
        next_row = len(names) + 1
        worksheet.update_cell(next_row, 1, employee_name)
        return next_row

    # ---------------------------------------------------------------------
    # Google Sheets formatting helpers
    # ---------------------------------------------------------------------
    @with_retry
    def _apply_background(self, worksheet, cell_range: str, color: Color):
        fmt = CellFormat(backgroundColor=color)
        format_cell_range(worksheet, cell_range, fmt)

    # ---------------------------------------------------------------------
    # Internal helpers to fill a single day
    # ---------------------------------------------------------------------
    def _fill_day(self, worksheet, day: datetime.date):
        """Fetch WebWork data for *day* and write it into the sheet."""

        # Skip weekends
        if day.weekday() >= 5:
            return

        print(f"  ↳ Filling {day.strftime('%A %d %b')} …")

        data = self.tracker.fetch_webwork_data(day.strftime("%Y-%m-%d"))
        _, first_entries, hr_team_emails = self.tracker.get_joining_times(data, target_date=day)

        col_idx = day.weekday() + 2  # Monday col 2 …
        start_dt = datetime.datetime.combine(day, datetime.datetime.strptime(START_TIME, "%H:%M").time(), tzinfo=self.tz)

        # Collect values and formatting first so we can batch the writes – this avoids
        # blowing past the Sheets 60-writes/min quota.
        cell_updates: list[gspread.Cell] = []
        fmt_tasks: list[tuple[str, Color]] = []  # (A1 notation, color)

        for email in hr_team_emails:
            name = self.tracker.get_user_name(email)
            row = self._ensure_employee_row(worksheet, name)
            cell_a1 = gspread.utils.rowcol_to_a1(row, col_idx)

            first_entry = first_entries.get(email)
            if not first_entry:
                value = "Absent"
                colour = self.RED
            else:
                minutes_late = (first_entry - start_dt).total_seconds() / 60.0
                time_str = first_entry.strftime("%I:%M %p")  # e.g., 12:15 AM
                value = f"'{time_str}"  # leading apostrophe → keep as text
                colour = self.YELLOW if minutes_late >= 5 else self.WHITE

            # Stage the value update and the formatting change.
            cell_updates.append(gspread.Cell(row, col_idx, value))
            fmt_tasks.append((cell_a1, colour))

        # --- Push all value updates in one request (with retry)
        @with_retry
        def batch_update_cells():
            worksheet.update_cells(cell_updates, value_input_option="USER_ENTERED")

        if cell_updates:
            batch_update_cells()

        # --- Push all formatting updates in a single batch request
        if fmt_tasks:
            try:
                @with_retry
                def apply_all_formatting():
                    with batch_updater(worksheet.spreadsheet) as batch:
                        for cell_a1, colour in fmt_tasks:
                            self._apply_background(worksheet, cell_a1, colour)

                apply_all_formatting()
            except APIError as e:
                # Ignore benign "Must specify at least one request" when no formatting changes are necessary
                if "Must specify at least one request" not in str(e):
                    raise

    # ---------------------------------------------------------------------
    # Core public API
    # ---------------------------------------------------------------------
    def update_until_date(self, date: datetime.date):
        """Update all weekdays from Monday up to *date* (inclusive)."""

        if date.weekday() >= 5:
            print("Weekend – nothing to do.")
            return

        # Authenticate WebWork users once
        self.tracker.fetch_user_info()

        monday, friday = self._week_bounds(date)
        worksheet = self._get_or_create_week_sheet(monday, friday)

        current = monday
        while current <= date:
            self._fill_day(worksheet, current)
            current += datetime.timedelta(days=1)

        print("Week sheet updated successfully.")


if __name__ == "__main__":
    eastern = pytz.timezone("US/Eastern")
    today_eastern = datetime.datetime.now(eastern).date()
    updater = WeeklyAttendance()
    updater.update_until_date(today_eastern) 