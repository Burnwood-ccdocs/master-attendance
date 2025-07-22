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


def run_complete_execution():
    """
    Run the complete attendance reporting system automatically.
    This executes all morning reports and the end-of-day report.
    """
    print("=" * 80)
    print("ATTENDANCE TRACKER - COMPLETE AUTOMATIC EXECUTION")
    print("=" * 80)
    
    # Get current date and time
    now = datetime.datetime.now(pytz.timezone('America/New_York'))
    today = now.date()
    current_time = now.time()
    
    print(f"Execution Date: {today.strftime('%A, %Y-%m-%d')}")
    print(f"Execution Time: {current_time.strftime('%I:%M %p')} EST")
    print(f"Total Departments: {len(DEPARTMENTS_CONFIG)}")
    print(f"Department Groups: {len(REPORT_GROUPS)}")
    
    # Show department groups
    print("\nDepartment Groups:")
    for time_str, depts in REPORT_GROUPS.items():
        print(f"  {time_str}: {', '.join(depts)}")
    
    print("\n" + "=" * 80)
    print("STARTING EXECUTION...")
    print("=" * 80)
    
    # Step 1: Run all morning reports
    print("\n" + "=" * 60)
    print("STEP 1: RUNNING ALL MORNING REPORTS")
    print("=" * 60)
    
    morning_success_count = 0
    morning_total_count = len(REPORT_GROUPS)
    
    for report_time_str, depts in REPORT_GROUPS.items():
        try:
            print(f"\n--- Running morning report for {len(depts)} department(s) scheduled at {report_time_str} ---")
            run_department_group_report(depts, 'morning')
            print(f"--- Completed morning report for {', '.join(depts)} ---")
            morning_success_count += 1
        except Exception as e:
            print(f"--- ERROR in morning report for {', '.join(depts)}: {e} ---")
    
    print(f"\nMorning Reports: {morning_success_count}/{morning_total_count} completed successfully")
    
    # Step 2: Run end-of-day report
    print("\n" + "=" * 60)
    print("STEP 2: RUNNING END-OF-DAY REPORT")
    print("=" * 60)
    
    try:
        all_depts = list(DEPARTMENTS_CONFIG.keys())
        print(f"\n--- Running EOD report for all {len(all_depts)} departments ---")
        run_department_group_report(all_depts, 'eod')
        print(f"--- Completed EOD report for all departments ---")
        eod_success = True
    except Exception as e:
        print(f"--- ERROR in EOD report: {e} ---")
        eod_success = False
    
    # Summary
    print("\n" + "=" * 80)
    print("EXECUTION SUMMARY")
    print("=" * 80)
    print(f"Date: {today.strftime('%A, %Y-%m-%d')}")
    print(f"Time: {current_time.strftime('%I:%M %p')} EST")
    print(f"Morning Reports: {morning_success_count}/{morning_total_count} successful")
    print(f"End-of-Day Report: {'✓ Success' if eod_success else '✗ Failed'}")
    print(f"Total Departments Processed: {len(DEPARTMENTS_CONFIG)}")
    print("=" * 80)
    print("EXECUTION COMPLETE")
    print("=" * 80)


def run_morning_only():
    """Run only the morning reports."""
    print("=" * 60)
    print("RUNNING MORNING REPORTS ONLY")
    print("=" * 60)
    
    for report_time_str, depts in REPORT_GROUPS.items():
        print(f"\n--- Running morning report for {len(depts)} department(s) scheduled at {report_time_str} ---")
        run_department_group_report(depts, 'morning')
        print(f"--- Completed morning report for {', '.join(depts)} ---")


def run_eod_only():
    """Run only the end-of-day report."""
    print("=" * 60)
    print("RUNNING END-OF-DAY REPORT ONLY")
    print("=" * 60)
    
    all_depts = list(DEPARTMENTS_CONFIG.keys())
    print(f"\n--- Running EOD report for all {len(all_depts)} departments ---")
    run_department_group_report(all_depts, 'eod')
    print(f"--- Completed EOD report for all departments ---")


def main():
    """
    Main function with command line arguments for different execution modes.
    """
    import sys
    
    if len(sys.argv) > 1:
        mode = sys.argv[1].lower()
        
        if mode == 'morning':
            run_morning_only()
        elif mode == 'eod':
            run_eod_only()
        elif mode == 'complete':
            run_complete_execution()
        else:
            print("Usage: python main_auto_run.py [morning|eod|complete]")
            print("  morning  - Run only morning reports")
            print("  eod      - Run only end-of-day report")
            print("  complete - Run both morning and EOD reports (default)")
            sys.exit(1)
    else:
        # Default: run complete execution
        run_complete_execution()


if __name__ == "__main__":
    main() 