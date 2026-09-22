"""
Automated Backup Scheduler
Runs backups every 6 hours to ensure data is never lost
Can be run locally or on any server
"""
import os
import time
import schedule
import datetime
import shutil
from backup_database import create_backup, export_data_to_json, cleanup_old_backups
from import_attendance_from_sheets import import_attendance_from_sheet

def scheduled_attendance_import():
    """Run scheduled attendance import"""
    print(f"\n⏰ Scheduled Attendance Import at {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    try:
        import_attendance_from_sheet()
        print("✅ Attendance import successful")
    except Exception as e:
        print(f"❌ Attendance import failed: {e}")

def scheduled_backup():
    """Run scheduled backup"""
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"\n{'='*60}")
    print(f"⏰ Scheduled backup at {timestamp}")
    print(f"{'='*60}")
    
    try:
        backup_file = create_backup()
        if backup_file:
            print(f"✅ Backup successful: {backup_file}")
        else:
            print("❌ Backup failed!")
        
        # Export to JSON weekly (on Monday)
        if datetime.datetime.now().weekday() == 0:  # Monday
            print("\n📦 Weekly JSON export...")
            export_data_to_json()
        
        # Cleanup old backups monthly (keep 90 days)
        cleanup_old_backups(keep_days=90)
        
    except Exception as e:
        print(f"❌ Error during backup: {e}")
    
    print(f"{'='*60}\n")

def run_scheduler():
    """Run the backup scheduler"""
    print("🚀 Starting automated backup scheduler")
    print("📅 Backup schedule: Every 6 hours")
    print("📦 JSON export: Every Monday")
    print("🗑️ Cleanup: Monthly (keep 90 days)")
    print("\nPress Ctrl+C to stop\n")
    
    # Run backup immediately on start
    scheduled_backup()
    
    # Schedule backups every 6 hours
    schedule.every(6).hours.do(scheduled_backup)
    
    # Schedule attendance import every 5 seconds
    schedule.every(5).seconds.do(scheduled_attendance_import)
    
    # Keep running
    while True:
        schedule.run_pending()
        time.sleep(1)  # Check every second

if __name__ == '__main__':
    run_scheduler()
