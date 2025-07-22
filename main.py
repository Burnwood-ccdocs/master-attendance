import datetime
import pytz
from apscheduler.schedulers.blocking import BlockingScheduler
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


def main():
    """
    Schedules and runs the consolidated, group-based attendance reporting jobs.
    """
    scheduler = BlockingScheduler(timezone='America/New_York')
    
    print("--- Setting up consolidated morning schedules ---")
    for report_time_str, depts in REPORT_GROUPS.items():
        hour, minute = map(int, report_time_str.split(':'))
        scheduler.add_job(
            run_department_group_report, 
            'cron', 
            hour=hour, 
            minute=minute,
            args=[depts, 'morning']
        )
        print(f"  • Scheduled morning report for {len(depts)} department(s) at {report_time_str} EST.")

    # Schedule the single end-of-day job for ALL departments
    all_depts = list(DEPARTMENTS_CONFIG.keys())
    scheduler.add_job(run_department_group_report, 'cron', hour=17, minute=30, args=[all_depts, 'eod'])
    print("\n--- Scheduled consolidated EOD full report for 5:30 PM EST ---")
    
    print("\nScheduler is running. Press Ctrl+C to exit.")

    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        print("Scheduler stopped.")

if __name__ == "__main__":
    main() 