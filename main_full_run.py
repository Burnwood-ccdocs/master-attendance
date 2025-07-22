import datetime
import pytz
from generate_report import AttendanceReportGenerator
from config import DEPARTMENT_START_TIMES, DEPARTMENTS_CONFIG
import collections

# --- Reporting Group Definitions ---
# Group departments by their scheduled report time (30 mins after start time)
REPORT_GROUPS = collections.defaultdict(list)
for dept, start_time in DEPARTMENT_START_TIMES.items():
    hour, minute = map(int, start_time.split(':'))
    report_dt = datetime.datetime.combine(datetime.date.today(), datetime.time(hour, minute)) + datetime.timedelta(minutes=30)
    REPORT_GROUPS[report_dt.strftime('%H:%M')].append(dept)


def run_department_group_report(departments, run_type):
    """
    This is the new target for the scheduler. It runs the consolidated
    report for a specific list of departments.
    """
    print(f"*** Kicking off {run_type.upper()} report for group: {', '.join(departments)} ***")
    today = datetime.datetime.now(pytz.timezone('America/New_York')).date()
    reporter = AttendanceReportGenerator()
    reporter.run_consolidated_report(departments, today, run_type)
    print(f"*** Finished {run_type.upper()} report for group ***")


def run_all_morning_reports():
    """Run all morning reports for all department groups."""
    print("=" * 60)
    print("RUNNING ALL MORNING REPORTS")
    print("=" * 60)
    
    for report_time_str, depts in REPORT_GROUPS.items():
        print(f"\n--- Running morning report for {len(depts)} department(s) scheduled at {report_time_str} ---")
        run_department_group_report(depts, 'morning')
        print(f"--- Completed morning report for {', '.join(depts)} ---")


def run_end_of_day_report():
    """Run the end-of-day report for all departments."""
    print("=" * 60)
    print("RUNNING END-OF-DAY REPORT")
    print("=" * 60)
    
    all_depts = list(DEPARTMENTS_CONFIG.keys())
    print(f"\n--- Running EOD report for all {len(all_depts)} departments ---")
    run_department_group_report(all_depts, 'eod')
    print(f"--- Completed EOD report for all departments ---")


def run_single_department_test():
    """Run a test for a single department to verify functionality."""
    print("=" * 60)
    print("RUNNING SINGLE DEPARTMENT TEST")
    print("=" * 60)
    
    # Test with IT Dept since it has a defined start time
    test_dept = "IT Dept"
    print(f"\n--- Testing single department: {test_dept} ---")
    run_department_group_report([test_dept], 'morning')
    print(f"--- Completed single department test for {test_dept} ---")


def main():
    """
    Run the complete attendance reporting system without scheduler.
    This allows for immediate testing of all functionality.
    """
    print("Attendance Tracker - Full Execution Mode")
    print("=" * 60)
    
    # Get current date and time
    now = datetime.datetime.now(pytz.timezone('America/New_York'))
    today = now.date()
    current_time = now.time()
    
    print(f"Current Date: {today.strftime('%A, %Y-%m-%d')}")
    print(f"Current Time: {current_time.strftime('%I:%M %p')} EST")
    print(f"Total Departments: {len(DEPARTMENTS_CONFIG)}")
    print(f"Department Groups: {len(REPORT_GROUPS)}")
    
    # Show department groups
    print("\nDepartment Groups:")
    for time_str, depts in REPORT_GROUPS.items():
        print(f"  {time_str}: {', '.join(depts)}")
    
    print("\n" + "=" * 60)
    print("SELECT EXECUTION MODE:")
    print("1. Run all morning reports")
    print("2. Run end-of-day report")
    print("3. Run single department test")
    print("4. Run both morning and EOD reports")
    print("5. Exit")
    print("=" * 60)
    
    while True:
        try:
            choice = input("\nEnter your choice (1-5): ").strip()
            
            if choice == '1':
                run_all_morning_reports()
                break
            elif choice == '2':
                run_end_of_day_report()
                break
            elif choice == '3':
                run_single_department_test()
                break
            elif choice == '4':
                run_all_morning_reports()
                print("\n" + "=" * 60)
                run_end_of_day_report()
                break
            elif choice == '5':
                print("Exiting...")
                break
            else:
                print("Invalid choice. Please enter 1, 2, 3, 4, or 5.")
        except KeyboardInterrupt:
            print("\n\nExiting due to user interruption...")
            break
        except Exception as e:
            print(f"Error: {e}")
            print("Please try again.")
    
    print("\n" + "=" * 60)
    print("EXECUTION COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main() 