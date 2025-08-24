"""
Shared utilities for CalendarBridge
Centralizes common functionality to reduce code duplication
"""
import os
import json
import logging
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

class GoogleCalendarAuth:
    """Centralized Google Calendar authentication."""
    
    @staticmethod
    def get_service(token_file, credentials_file, scopes):
        """Returns authenticated Google Calendar service."""
        creds = None
        if os.path.exists(token_file):
            creds = Credentials.from_authorized_user_file(token_file, scopes)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    logging.error(f"Error refreshing token: {e}. Re-authenticating.")
                    flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
                creds = flow.run_local_server(port=0)
            with open(token_file, 'w') as token:
                token.write(creds.to_json())
        return build('calendar', 'v3', credentials=creds)

class ConfigManager:
    """Centralized configuration management."""
    
    @staticmethod
    def load_config(config_file):
        """Load and validate configuration."""
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        # Validate required fields
        required_fields = ['outlook_calendar_name', 'google_calendar_id']
        for field in required_fields:
            if field not in config:
                raise ValueError(f"Missing required config field: {field}")
        
        # Set defaults for optional fields
        defaults = {
            'outlook_calendar_index': 1,
            'sync_days_past': 90,
            'sync_days_future': 120,
            'api_delay_seconds': 0.1,
            'timezone': 'America/New_York',
            'enable_batch_operations': False,
            'batch_size': 50,
            'enable_health_monitoring': True,
            'enable_state_tracking': True
        }
        
        for key, value in defaults.items():
            if key not in config:
                config[key] = value
        
        return config
