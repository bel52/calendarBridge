#!/usr/bin/env python3
"""
Test script to verify enhanced CalendarBridge functionality
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))

from shared_utils import ConfigManager, GoogleCalendarAuth

def test_config():
    """Test configuration loading and validation."""
    print("Testing configuration...")
    try:
        config = ConfigManager.load_config('calendar_config.json')
        print("✓ Configuration loaded successfully")
        print(f"  Calendar ID: {config.get('google_calendar_id')}")
        print(f"  Timezone: {config.get('timezone')}")
        print(f"  Batch operations: {'Enabled' if config.get('enable_batch_operations') else 'Disabled'}")
        print(f"  Health monitoring: {'Enabled' if config.get('enable_health_monitoring') else 'Disabled'}")
        return True
    except Exception as e:
        print(f"✗ Configuration test failed: {e}")
        return False

def test_auth():
    """Test Google authentication."""
    print("\nTesting Google authentication...")
    try:
        service = GoogleCalendarAuth.get_service(
            'token.json', 
            'credentials.json',
            ['https://www.googleapis.com/auth/calendar']
        )
        # Try to list calendars
        result = service.calendarList().list(maxResults=1).execute()
        print("✓ Google authentication successful")
        return True
    except Exception as e:
        print(f"✗ Authentication test failed: {e}")
        return False

def main():
    print("="*50)
    print("CalendarBridge Enhanced Setup Test")
    print("="*50)
    
    all_tests_passed = True
    
    # Run tests
    if not test_config():
        all_tests_passed = False
    
    if not test_auth():
        all_tests_passed = False
    
    print("\n" + "="*50)
    if all_tests_passed:
        print("✅ All tests passed! Your enhanced setup is ready.")
    else:
        print("⚠️  Some tests failed. Please check the errors above.")
    print("="*50)

if __name__ == '__main__':
    main()
