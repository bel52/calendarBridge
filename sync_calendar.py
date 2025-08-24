import asyncio
from datetime import datetime, timedelta
import configparser

from azure.identity.aio import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

# Load configuration from a file
config = configparser.ConfigParser()
config.read(['config.cfg'])
CLIENT_ID = config['azure']['ClientId']
TENANT_ID = config['azure']
GRAPH_SCOPES = ['https://graph.microsoft.com/.default']

async def main():
    """Main function to authenticate and fetch calendar events."""
    print("Python Microsoft Graph Calendar Sync")
    print("-" * 35)

    try:
        # Create a credential object for the device code flow.
        # This will prompt the user to sign in using a code in their browser.
        credential = DeviceCodeCredential(client_id=CLIENT_ID, tenant_id=TENANT_ID)

        # Initialize the Graph Service Client with the credential
        graph_client = GraphServiceClient(credentials=credential, scopes=GRAPH_SCOPES)

        # --- Define the time range for the calendar view ---
        # Get events from today for the next 7 days
        start_time = datetime.now()
        end_time = start_time + timedelta(days=7)

        # Format dates in ISO 8601 format required by the API
        start_iso = start_time.isoformat()
        end_iso = end_time.isoformat()

        print(f"\nFetching calendar events from {start_time.strftime('%Y-%m-%d')} to {end_time.strftime('%Y-%m-%d')}...\n")

        # --- Make the API call to get the calendar view ---
        # The calendarView endpoint is efficient for fetching events in a specific window.
        # It correctly expands recurring events.
        # See: https://learn.microsoft.com/en-us/graph/api/user-list-calendarview
        events_response = await graph_client.me.calendar_view.get(
            query_parameters={
                "startDateTime": start_iso,
                "endDateTime": end_iso,
                "select": ["subject", "start", "end", "organizer"],
                "orderby":,
            }
        )

        if events_response and events_response.value:
            print(f"Found {len(events_response.value)} events:")
            for event in events_response.value:
                # The API returns times in UTC by default.
                # Here we just print them, but they could be converted to local time.
                start_event_time = datetime.fromisoformat(event.start.date_time)
                organizer_name = event.organizer.email_address.name if event.organizer and event.organizer.email_address else "N/A"
                
                print(f"  - Subject: {event.subject}")
                print(f"    Start:   {start_event_time.strftime('%Y-%m-%d %H:%M')}")
                print(f"    Organizer: {organizer_name}\n")
        else:
            print("No events found in the specified time range.")

    except ODataError as odata_error:
        print(f"Graph API Error: {odata_error.error.code}")
        print(f"Message: {odata_error.error.message}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    asyncio.run(main())
