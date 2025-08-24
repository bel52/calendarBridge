#!/usr/bin/env python3
"""
CalendarBridge Monitoring Tool
Checks sync health and displays statistics
"""
import json
import os
import datetime
from pathlib import Path

APP_DIR = os.path.dirname(os.path.realpath(__file__))
STATE_FILE = os.path.join(APP_DIR, 'sync_state.json')
HEALTH_FILE = '/tmp/calendarbridge_health.json'
LOG_FILE = os.path.join(APP_DIR, 'logs', 'calendar_sync.log')

def format_time_ago(iso_timestamp):
    """Convert ISO timestamp to human-readable time ago."""
    if not iso_timestamp:
        return "Never"
    
    dt = datetime.datetime.fromisoformat(iso_timestamp.replace('Z', '+00:00'))
    now = datetime.datetime.now(datetime.timezone.utc)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.timezone.utc)
    
    delta = now - dt
    if delta.days > 0:
        return f"{delta.days} day{'s' if delta.days != 1 else ''} ago"
    elif delta.seconds > 3600:
        hours = delta.seconds // 3600
        return f"{hours} hour{'s' if hours != 1 else ''} ago"
    elif delta.seconds > 60:
        minutes = delta.seconds // 60
        return f"{minutes} minute{'s' if minutes != 1 else ''} ago"
    else:
        return "Just now"

def main():
    print("="*60)
    print("CalendarBridge Sync Monitor".center(60))
    print("="*60)
    
    # Check health status
    if os.path.exists(HEALTH_FILE):
        with open(HEALTH_FILE, 'r') as f:
            health = json.load(f)
        
        status_emoji = "✅" if health['status'] == 'healthy' else "⚠️"
        print(f"\nHealth Status: {status_emoji} {health['status'].upper()}")
        print(f"Version: {health.get('version', 'Unknown')}")
        
        if health.get('last_successful_sync'):
            print(f"Last Successful Sync: {format_time_ago(health['last_successful_sync'])}")
        
        if health.get('consecutive_failures', 0) > 0:
            print(f"⚠️  Consecutive Failures: {health['consecutive_failures']}")
        
        if health.get('last_error'):
            print(f"Last Error: {health['last_error'][:100]}")
    else:
        print("\n⚠️  No health file found. Sync may not have run yet.")
    
    # Check sync state
    if os.path.exists(STATE_FILE):
        print("\n" + "-"*60)
        print("Sync Statistics".center(60))
        print("-"*60)
        
        with open(STATE_FILE, 'r') as f:
            state = json.load(f)
        
        print(f"Last Run: {format_time_ago(state.get('last_sync'))}")
        print(f"Duration: {state.get('sync_duration_seconds', 0):.1f} seconds")
        print(f"\nEvents:")
        print(f"  • Created: {state.get('events_created', 0)}")
        print(f"  • Updated: {state.get('events_updated', 0)}")
        print(f"  • Deleted: {state.get('events_deleted', 0)}")
        
        if state.get('errors'):
            print(f"\n⚠️  Recent Errors ({len(state['errors'])})")
            for error in state['errors'][:3]:
                print(f"  • {error[:100]}")
    else:
        print("\n⚠️  No state file found. Sync may not have run yet.")
    
    # Check log file
    if os.path.exists(LOG_FILE):
        log_size = Path(LOG_FILE).stat().st_size / (1024 * 1024)  # MB
        print(f"\n📄 Log File: {log_size:.1f} MB")
        
        # Show last few log lines
        print("\nRecent Log Entries:")
        try:
            with open(LOG_FILE, 'r') as f:
                lines = f.readlines()
                for line in lines[-5:]:
                    print(f"  {line.strip()[:100]}")
        except Exception as e:
            print(f"  Could not read log: {e}")
    
    print("\n" + "="*60)
    print("Monitoring complete. Check logs for detailed information.")
    print("="*60)

if __name__ == '__main__':
    main()
