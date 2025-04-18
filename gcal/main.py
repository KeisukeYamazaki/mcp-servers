import json
import logging
import os
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from mcp.server.fastmcp import Context, FastMCP

# Log configuration
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Set required scopes for Google APIs
SCOPES = ["https://www.googleapis.com/auth/calendar"]

# Create MCP server instance
mcp = FastMCP("GoogleCalendarServer")


def get_credentials():
    """Get authentication credentials for connecting to Google API"""
    creds = None
    token_path = "token.json"
    credentials_path = "credentials.json"

    # Check if credentials.json exists
    if not os.path.exists(credentials_path):
        logger.error(
            f"Authentication file '{credentials_path}' not found. Please download OAuth client ID from Google Cloud Console."
        )
        raise FileNotFoundError(f"Authentication file '{credentials_path}' not found")

    # token.json stores authenticated user information
    logger.info("Checking authentication credentials...")
    if os.path.exists(token_path):
        logger.info(f"Loading token file '{token_path}'...")
        try:
            with open(token_path, "r") as token_file:
                token_data = json.load(token_file)
                creds = Credentials.from_authorized_user_info(token_data, SCOPES)
                logger.info("Existing credentials loaded")
        except Exception as e:
            logger.error(f"Token file loading error: {e}")
            creds = None

    # Update credentials if expired or non-existent
    if not creds or not creds.valid:
        logger.info("Credentials are invalid or expired. Updating...")
        if creds and creds.expired and creds.refresh_token:
            logger.info("Refreshing token...")
            try:
                creds.refresh(Request())
                logger.info("Credentials updated")
            except Exception as e:
                logger.error(f"Failed to refresh token: {e}")
                creds = None

        if not creds:
            logger.info("Starting new authentication flow. Browser will open...")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)
                logger.info("New credentials acquired")
            except Exception as e:
                logger.error(f"Failed to execute authentication flow: {e}")
                raise

        # Save credentials for next time
        logger.info(f"Saving credentials to '{token_path}'...")
        try:
            with open(token_path, "w") as token:
                token.write(creds.to_json())
            logger.info("Credentials saved")
        except Exception as e:
            logger.error(f"Failed to save token: {e}")

    return creds


def format_event(event, calendar_name=None):
    """Format calendar event data for better readability"""
    start = event.get("start", {}).get("dateTime", event.get("start", {}).get("date", "Not set"))
    end = event.get("end", {}).get("dateTime", event.get("end", {}).get("date", "Not set"))

    # Format datetime if needed
    if "T" in start:
        start_dt = datetime.fromisoformat(start.replace("Z", "+00:00"))
        start = start_dt.strftime("%Y-%m-%d %H:%M")
    if "T" in end:
        end_dt = datetime.fromisoformat(end.replace("Z", "+00:00"))
        end = end_dt.strftime("%Y-%m-%d %H:%M")

    location = event.get("location", "Not set")
    description = event.get("description", "")

    formatted = f"Event: {event.get('summary', 'Untitled Event')}\n"
    if calendar_name:
        formatted += f"Calendar: {calendar_name}\n"
    formatted += f"Start: {start}\n"
    formatted += f"End: {end}\n"

    if location != "Not set":
        formatted += f"Location: {location}\n"

    if description:
        formatted += f"Details: {description}\n"

    formatted += f"ID: {event.get('id')}\n"
    if calendar_name:
        formatted += f"Calendar ID: {event.get('organizer', {}).get('email', '')}\n"

    return formatted


def get_all_calendars(service):
    """Get a list of all available calendars"""
    calendar_list = service.calendarList().list().execute()
    return calendar_list.get("items", [])


@mcp.tool()
def list_upcoming_events(max_results: int = 10, days_ahead: int = 7, use_all_calendars: bool = True) -> str:
    """
    Get upcoming calendar events

    Args:
        max_results: Maximum number of events to retrieve (default: 10)
        days_ahead: Number of days ahead to retrieve events (default: 7)
        use_all_calendars: Whether to get events from all calendars (default: True)

    Returns:
        List of retrieved events
    """
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get current time and time 'days_ahead' days in the future
        now = datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
        future = (datetime.utcnow() + timedelta(days=days_ahead)).isoformat() + "Z"

        all_events = []
        calendar_names = {}

        if use_all_calendars:
            # Get all calendars
            calendars = get_all_calendars(service)

            # If no calendars found, fall back to primary
            if not calendars:
                calendars = [{"id": "primary", "summary": "Primary"}]
        else:
            # Just use primary calendar
            calendars = [{"id": "primary", "summary": "Primary"}]

        # Store calendar names for reference
        for calendar in calendars:
            calendar_names[calendar["id"]] = calendar.get("summary", calendar["id"])

        # Get events from each calendar
        for calendar in calendars:
            calendar_id = calendar["id"]
            try:
                events_result = (
                    service.events()
                    .list(
                        calendarId=calendar_id,
                        timeMin=now,
                        timeMax=future,
                        maxResults=max_results,
                        singleEvents=True,
                        orderBy="startTime",
                    )
                    .execute()
                )

                events = events_result.get("items", [])
                for event in events:
                    event["_calendar_name"] = calendar_names.get(calendar_id, calendar_id)
                all_events.extend(events)
            except HttpError as e:
                logger.error(f"Error getting events for calendar {calendar_id}: {e}")

        # Sort all events by start time
        all_events.sort(key=lambda x: x.get("start", {}).get("dateTime", x.get("start", {}).get("date", "")))

        # Limit to max_results
        all_events = all_events[:max_results]

        if not all_events:
            return "No upcoming events."

        result = f"Next {len(all_events)} events:\n\n"
        for event in all_events:
            calendar_name = event.pop("_calendar_name", None)
            result += format_event(event, calendar_name) + "\n" + "-" * 40 + "\n"

        return result

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def get_event_details(event_id: str, calendar_id: str = "primary") -> str:
    """
    Get details for a specific event

    Args:
        event_id: Event ID
        calendar_id: Calendar ID (default: "primary")

    Returns:
        Detailed information about the event
    """
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get calendar name
        calendar_name = None
        if calendar_id != "primary":
            try:
                calendar_info = service.calendarList().get(calendarId=calendar_id).execute()
                calendar_name = calendar_info.get("summary", calendar_id)
            except HttpError:
                calendar_name = calendar_id

        # Get event details
        event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()

        if not event:
            return f"Event not found: {event_id}"

        return format_event(event, calendar_name)

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def create_event(
    summary: str,
    start_datetime: str,
    end_datetime: str,
    description: str = "",
    location: str = "",
    attendees: str = "",
    calendar_id: str = "primary",
) -> str:
    """
    Create a new calendar event

    Args:
        summary: Event title
        start_datetime: Start date and time (format YYYY-MM-DDTHH:MM:SS e.g., 2024-05-20T10:00:00)
        end_datetime: End date and time (format YYYY-MM-DDTHH:MM:SS e.g., 2024-05-20T11:00:00)
        description: Event description (optional)
        location: Event location (optional)
        attendees: Attendee email addresses (comma-separated, optional)
        calendar_id: Calendar ID to create the event in (default: "primary")

    Returns:
        Information about the created event
    """
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get calendar name
        calendar_name = None
        if calendar_id != "primary":
            try:
                calendar_info = service.calendarList().get(calendarId=calendar_id).execute()
                calendar_name = calendar_info.get("summary", calendar_id)
            except HttpError:
                calendar_name = calendar_id

        # Parse attendees
        attendee_list = []
        if attendees:
            for email in attendees.split(","):
                email = email.strip()
                if email:
                    attendee_list.append({"email": email})

        # Create event
        event = {
            "summary": summary,
            "location": location,
            "description": description,
            "start": {
                "dateTime": start_datetime,
                "timeZone": "Asia/Tokyo",
            },
            "end": {
                "dateTime": end_datetime,
                "timeZone": "Asia/Tokyo",
            },
        }

        if attendee_list:
            event["attendees"] = attendee_list

        event = service.events().insert(calendarId=calendar_id, body=event).execute()

        return f"Event created: {event.get('htmlLink')}\n\n" + format_event(event, calendar_name)

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def update_event(
    event_id: str,
    calendar_id: str = "primary",
    summary: str = None,
    start_datetime: str = None,
    end_datetime: str = None,
    description: str = None,
    location: str = None,
) -> str:
    """
    Update an existing calendar event

    Args:
        event_id: ID of the event to update
        calendar_id: Calendar ID (default: "primary")
        summary: New event title (optional)
        start_datetime: New start date and time (format YYYY-MM-DDTHH:MM:SS e.g., 2024-05-20T10:00:00, optional)
        end_datetime: New end date and time (format YYYY-MM-DDTHH:MM:SS e.g., 2024-05-20T11:00:00, optional)
        description: New event description (optional)
        location: New event location (optional)

    Returns:
        Information about the updated event
    """
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get calendar name
        calendar_name = None
        if calendar_id != "primary":
            try:
                calendar_info = service.calendarList().get(calendarId=calendar_id).execute()
                calendar_name = calendar_info.get("summary", calendar_id)
            except HttpError:
                calendar_name = calendar_id

        # Get current event
        event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()

        # Update fields if provided
        if summary:
            event["summary"] = summary

        if description is not None:
            event["description"] = description

        if location is not None:
            event["location"] = location

        if start_datetime:
            event["start"]["dateTime"] = start_datetime

        if end_datetime:
            event["end"]["dateTime"] = end_datetime

        # Update event
        updated_event = service.events().update(calendarId=calendar_id, eventId=event_id, body=event).execute()

        return f"Event updated:\n\n" + format_event(updated_event, calendar_name)

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def delete_event(event_id: str, calendar_id: str = "primary") -> str:
    """
    Delete a calendar event

    Args:
        event_id: ID of the event to delete
        calendar_id: Calendar ID (default: "primary")

    Returns:
        Result of the delete operation
    """
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get calendar name
        calendar_name = None
        if calendar_id != "primary":
            try:
                calendar_info = service.calendarList().get(calendarId=calendar_id).execute()
                calendar_name = calendar_info.get("summary", calendar_id)
            except HttpError:
                pass

        # Get event details first for confirmation
        event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        summary = event.get("summary", "Untitled Event")

        # Delete event
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()

        calendar_info = f" (Calendar: {calendar_name})" if calendar_name else ""
        return f"Event '{summary}' (ID: {event_id}){calendar_info} has been deleted."

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def search_events(query: str, max_results: int = 10, use_all_calendars: bool = True) -> str:
    """
    Search for calendar events

    Args:
        query: Search keyword
        max_results: Maximum number of search results (default: 10)
        use_all_calendars: Whether to search in all calendars (default: True)

    Returns:
        List of events matching the search query
    """
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get current time
        now = datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time

        all_events = []
        calendar_names = {}

        if use_all_calendars:
            # Get all calendars
            calendars = get_all_calendars(service)

            # If no calendars found, fall back to primary
            if not calendars:
                calendars = [{"id": "primary", "summary": "Primary"}]
        else:
            # Just use primary calendar
            calendars = [{"id": "primary", "summary": "Primary"}]

        # Store calendar names for reference
        for calendar in calendars:
            calendar_names[calendar["id"]] = calendar.get("summary", calendar["id"])

        # Search for events in each calendar
        for calendar in calendars:
            calendar_id = calendar["id"]
            try:
                events_result = (
                    service.events()
                    .list(
                        calendarId=calendar_id,
                        timeMin=now,
                        maxResults=max_results,
                        singleEvents=True,
                        orderBy="startTime",
                        q=query,
                    )
                    .execute()
                )

                events = events_result.get("items", [])
                for event in events:
                    event["_calendar_name"] = calendar_names.get(calendar_id, calendar_id)
                all_events.extend(events)
            except HttpError as e:
                logger.error(f"Error searching events for calendar {calendar_id}: {e}")

        # Sort all events by start time
        all_events.sort(key=lambda x: x.get("start", {}).get("dateTime", x.get("start", {}).get("date", "")))

        # Limit to max_results
        all_events = all_events[:max_results]

        if not all_events:
            return f"No events matching search query '{query}' were found."

        result = f"Search results for '{query}' ({len(all_events)} events):\n\n"
        for event in all_events:
            calendar_name = event.pop("_calendar_name", None)
            result += format_event(event, calendar_name) + "\n" + "-" * 40 + "\n"

        return result

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.resource("gcal://upcoming")
def upcoming_events_resource() -> str:
    """Provide upcoming events as a resource"""
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get current time and time 7 days in the future
        now = datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
        future = (datetime.utcnow() + timedelta(days=7)).isoformat() + "Z"

        all_events = []
        calendar_names = {}

        # Get all calendars
        calendars = get_all_calendars(service)

        # If no calendars found, fall back to primary
        if not calendars:
            calendars = [{"id": "primary", "summary": "Primary"}]

        # Store calendar names for reference
        for calendar in calendars:
            calendar_names[calendar["id"]] = calendar.get("summary", calendar["id"])

        # Get events from each calendar
        for calendar in calendars:
            calendar_id = calendar["id"]
            try:
                events_result = (
                    service.events()
                    .list(
                        calendarId=calendar_id,
                        timeMin=now,
                        timeMax=future,
                        maxResults=10,
                        singleEvents=True,
                        orderBy="startTime",
                    )
                    .execute()
                )

                events = events_result.get("items", [])
                for event in events:
                    event["_calendar_name"] = calendar_names.get(calendar_id, calendar_id)
                all_events.extend(events)
            except HttpError as e:
                logger.error(f"Error getting events for calendar {calendar_id}: {e}")

        # Sort all events by start time
        all_events.sort(key=lambda x: x.get("start", {}).get("dateTime", x.get("start", {}).get("date", "")))

        # Limit to 10 results
        all_events = all_events[:10]

        if not all_events:
            return "No upcoming events."

        result = f"Upcoming events:\n\n"
        for event in all_events:
            calendar_name = event.pop("_calendar_name", None)
            result += format_event(event, calendar_name) + "\n" + "-" * 40 + "\n"

        return result

    except Exception as e:
        return f"Resource retrieval error: {str(e)}"


@mcp.resource("gcal://event/{event_id}")
def event_resource(event_id: str) -> str:
    """Provide a specific event as a resource"""
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Try to find the event in any calendar
        calendars = get_all_calendars(service)

        # If no calendars found, fall back to primary only
        if not calendars:
            calendars = [{"id": "primary", "summary": "Primary"}]

        for calendar in calendars:
            calendar_id = calendar["id"]
            calendar_name = calendar.get("summary", calendar_id)

            try:
                # Get event details
                event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
                return format_event(event, calendar_name)
            except HttpError:
                # If not found, continue to next calendar
                continue

        return f"Event not found: {event_id}"

    except Exception as e:
        return f"Resource retrieval error: {str(e)}"


@mcp.resource("gcal://today")
def today_events_resource() -> str:
    """Provide today's events as a resource"""
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get today's start and end
        today = datetime.now().date()
        today_start = datetime.combine(today, datetime.min.time()).isoformat() + "Z"
        today_end = datetime.combine(today, datetime.max.time()).isoformat() + "Z"

        all_events = []
        calendar_names = {}

        # Get all calendars
        calendars = get_all_calendars(service)

        # If no calendars found, fall back to primary
        if not calendars:
            calendars = [{"id": "primary", "summary": "Primary"}]

        # Store calendar names for reference
        for calendar in calendars:
            calendar_names[calendar["id"]] = calendar.get("summary", calendar["id"])

        # Get events from each calendar
        for calendar in calendars:
            calendar_id = calendar["id"]
            try:
                events_result = (
                    service.events()
                    .list(
                        calendarId=calendar_id,
                        timeMin=today_start,
                        timeMax=today_end,
                        singleEvents=True,
                        orderBy="startTime",
                    )
                    .execute()
                )

                events = events_result.get("items", [])
                for event in events:
                    event["_calendar_name"] = calendar_names.get(calendar_id, calendar_id)
                all_events.extend(events)
            except HttpError as e:
                logger.error(f"Error getting events for calendar {calendar_id}: {e}")

        # Sort all events by start time
        all_events.sort(key=lambda x: x.get("start", {}).get("dateTime", x.get("start", {}).get("date", "")))

        if not all_events:
            return "No events today."

        result = f"Today's events ({len(all_events)}):\n\n"
        for event in all_events:
            calendar_name = event.pop("_calendar_name", None)
            result += format_event(event, calendar_name) + "\n" + "-" * 40 + "\n"

        return result

    except Exception as e:
        return f"Resource retrieval error: {str(e)}"


@mcp.tool()
def list_calendars() -> str:
    """
    Get a list of available calendars

    Returns:
        List of available calendars
    """
    try:
        # Connect to Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get calendar list
        calendar_list = service.calendarList().list().execute()
        calendars = calendar_list.get("items", [])

        if not calendars:
            return "No available calendars found."

        result = "Available calendars:\n\n"
        for calendar in calendars:
            result += f"Name: {calendar.get('summary')}\n"
            result += f"ID: {calendar.get('id')}\n"
            result += f"Description: {calendar.get('description', 'No description')}\n"
            result += "-" * 40 + "\n"

        return result

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def move_event(event_id: str, source_calendar_id: str, destination_calendar_id: str, send_updates: str = "none") -> str:
    """
    Move an event to a different calendar

    Args:
        event_id: ID of the event to move
        source_calendar_id: Source calendar ID
        destination_calendar_id: Destination calendar ID
        send_updates: How to send notifications ('all', 'externalOnly', or 'none', default: 'none')

    Returns:
        Result of the move operation
    """
    try:
        # Connect to Google Calendar API
        creds = get_credentials()
        service = build("calendar", "v3", credentials=creds)

        # Get source calendar name
        source_calendar_name = None
        if source_calendar_id != "primary":
            try:
                calendar_info = service.calendarList().get(calendarId=source_calendar_id).execute()
                source_calendar_name = calendar_info.get("summary", source_calendar_id)
            except HttpError:
                source_calendar_name = source_calendar_id

        # Get destination calendar name
        destination_calendar_name = None
        if destination_calendar_id != "primary":
            try:
                calendar_info = service.calendarList().get(calendarId=destination_calendar_id).execute()
                destination_calendar_name = calendar_info.get("summary", destination_calendar_id)
            except HttpError:
                destination_calendar_name = destination_calendar_id

        # Get event details first for confirmation
        event = service.events().get(calendarId=source_calendar_id, eventId=event_id).execute()
        summary = event.get("summary", "Untitled Event")

        # Move event to destination calendar
        moved_event = (
            service.events()
            .move(
                calendarId=source_calendar_id,
                eventId=event_id,
                destination=destination_calendar_id,
                sendUpdates=send_updates,
            )
            .execute()
        )

        # Format response
        from_calendar = f"'{source_calendar_name}'" if source_calendar_name else source_calendar_id
        to_calendar = f"'{destination_calendar_name}'" if destination_calendar_name else destination_calendar_id

        result = f"Event '{summary}' has been moved from {from_calendar} to {to_calendar}.\n\n"
        result += format_event(moved_event, destination_calendar_name)

        return result

    except HttpError as error:
        if "birthdayEvents, focusTime, fromGmail, outOfOffice and workingLocation events cannot be moved" in str(error):
            return "Error: Birthday, focus time, Gmail-derived, out-of-office, and working location events cannot be moved."
        return f"An error occurred: {error}"


if __name__ == "__main__":
    mcp.run()
