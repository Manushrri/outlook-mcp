"""
Microsoft Outlook Calendar Tools
"""

import base64
import os
from typing import Optional, Literal, List


def create_calendar(
    client,
    name: str,
    color: Optional[Literal["auto", "lightBlue", "lightGreen", "lightOrange", "lightGray", "lightYellow", "lightTeal", "lightPink", "lightBrown", "lightPurple", "lightRed"]] = None,
    hexColor: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Create a new calendar in the signed-in user's mailbox.
    Use when organizing events into a separate calendar.
    
    Args:
        client: The OutlookClient instance
        name: The name of the calendar
        color: Optional calendar color preset
        hexColor: Optional hex color code for the calendar
        user_id: Optional user ID (defaults to 'me')
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Build the calendar payload
        calendar_data = {
            "name": name
        }
        
        # Add optional fields if provided
        if color is not None:
            calendar_data["color"] = color
        if hexColor is not None:
            calendar_data["hexColor"] = hexColor
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/calendars"
        
        # Make the API call
        result = client.post(endpoint, json=calendar_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def add_event_attachment(
    client,
    event_id: str,
    name: str,
    odata_type: Literal["#microsoft.graph.fileAttachment", "#microsoft.graph.itemAttachment"],
    contentBytes: Optional[str] = None,
    file_path: Optional[str] = None,
    text_content: Optional[str] = None,
    item: Optional[dict] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Adds an attachment to a specific Outlook calendar event.
    Use when you need to attach a file or nested item to an existing event.
    
    For file attachments, you can provide content in three ways (in order of priority):
    1. file_path: Path to a file - will be read and Base64 encoded automatically
    2. text_content: Plain text - will be Base64 encoded automatically
    3. contentBytes: Already Base64-encoded content (manual)
    
    Args:
        client: The OutlookClient instance
        event_id: The ID of the event
        name: The name of the attachment (e.g., "document.pdf", "notes.txt")
        odata_type: The type of attachment ("#microsoft.graph.fileAttachment" or "#microsoft.graph.itemAttachment")
        contentBytes: Base64-encoded content (for file attachments) - use if you already have Base64
        file_path: Path to a file to attach - will be automatically Base64 encoded
        text_content: Plain text content to attach - will be automatically Base64 encoded
        item: Item data (for item attachments)
        user_id: Optional user ID (defaults to 'me')
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Build the attachment payload
        attachment_data = {
            "@odata.type": odata_type,
            "name": name
        }
        
        # Handle contentBytes - priority: file_path > text_content > contentBytes
        final_content_bytes = None
        
        if file_path is not None:
            # Read file and encode to Base64
            if not os.path.exists(file_path):
                return {
                    "successful": False,
                    "data": {},
                    "error": f"File not found: {file_path}"
                }
            try:
                with open(file_path, "rb") as f:
                    file_content = f.read()
                final_content_bytes = base64.b64encode(file_content).decode("utf-8")
            except Exception as file_error:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Error reading file: {str(file_error)}"
                }
        elif text_content is not None:
            # Encode plain text to Base64
            final_content_bytes = base64.b64encode(text_content.encode("utf-8")).decode("utf-8")
        elif contentBytes is not None:
            # Use provided Base64 content directly
            final_content_bytes = contentBytes
        
        if final_content_bytes is not None:
            attachment_data["contentBytes"] = final_content_bytes
        if item is not None:
            attachment_data["item"] = item
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/events/{event_id}/attachments"
        
        # Make the API call
        result = client.post(endpoint, json=attachment_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def create_event(
    client,
    subject: str,
    body: str,
    start_datetime: str,
    end_datetime: str,
    time_zone: str,
    location: Optional[str] = None,
    attendees_info: Optional[List[dict]] = None,
    categories: Optional[List[str]] = None,
    is_html: Optional[bool] = None,
    is_online_meeting: Optional[bool] = None,
    online_meeting_provider: Optional[str] = None,
    show_as: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Creates a new Outlook calendar event.
    Ensures start_datetime is chronologically before end_datetime.
    
    Args:
        client: The OutlookClient instance
        subject: The subject of the event
        body: The body content of the event
        start_datetime: Start date/time (ISO 8601 format)
        end_datetime: End date/time (ISO 8601 format)
        time_zone: Time zone for the event
        location: Optional location
        attendees_info: Optional list of attendee info dicts
        categories: Optional list of categories
        is_html: Whether body is HTML
        is_online_meeting: Whether it's an online meeting
        online_meeting_provider: Online meeting provider
        show_as: Show as status (free, tentative, busy, oof, workingElsewhere)
        user_id: Optional user ID (defaults to 'me')
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Build the event payload
        event_data = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body
            },
            "start": {
                "dateTime": start_datetime,
                "timeZone": time_zone
            },
            "end": {
                "dateTime": end_datetime,
                "timeZone": time_zone
            }
        }
        
        # Add optional fields
        if location is not None:
            event_data["location"] = {"displayName": location}
        
        if attendees_info is not None:
            event_data["attendees"] = [
                {
                    "emailAddress": attendee.get("emailAddress", {}),
                    "type": attendee.get("type", "required")
                }
                for attendee in attendees_info
            ]
        
        if categories is not None:
            event_data["categories"] = categories
        
        if is_online_meeting is not None:
            event_data["isOnlineMeeting"] = is_online_meeting
        
        if online_meeting_provider is not None:
            event_data["onlineMeetingProvider"] = online_meeting_provider
        
        if show_as is not None:
            event_data["showAs"] = show_as
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/events"
        
        # Make the API call
        result = client.post(endpoint, json=event_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def delete_event(
    client,
    event_id: str,
    send_notifications: Optional[bool] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Deletes an existing calendar event from the user's Outlook calendar.
    Optionally sends cancellation notifications to attendees.
    
    Args:
        client: The OutlookClient instance
        event_id: The ID of the event to delete
        send_notifications: Whether to send cancellation notifications to attendees
        user_id: Optional user ID (defaults to 'me')
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/events/{event_id}"
        
        # Add header for cancellation notifications if specified
        headers = {}
        if send_notifications is False:
            # Use the Prefer header to suppress notifications
            headers["Prefer"] = "outlook.notification-handling=suppress"
        
        # Make the API call
        client.delete(endpoint, headers=headers if headers else None)
        
        return {
            "successful": True,
            "data": {"message": "Event deleted successfully"}
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def get_event(
    client,
    event_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves the full details of a specific calendar event by its ID.
    
    Args:
        client: The OutlookClient instance
        event_id: The ID of the event to retrieve
        user_id: Optional user ID (defaults to 'me')
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/events/{event_id}"
        
        # Make the API call
        result = client.get(endpoint)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def update_calendar_event(
    client,
    event_id: str,
    subject: Optional[str] = None,
    body: Optional[dict] = None,
    start_datetime: Optional[str] = None,
    end_datetime: Optional[str] = None,
    time_zone: Optional[str] = None,
    location: Optional[dict] = None,
    attendees: Optional[List[dict]] = None,
    categories: Optional[List[str]] = None,
    show_as: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Updates specified fields of an existing Outlook calendar event.
    
    Args:
        client: The OutlookClient instance
        event_id: The ID of the event to update
        subject: Optional subject of the event
        body: Optional body object with contentType and content
        start_datetime: Optional start date/time (ISO 8601 format)
        end_datetime: Optional end date/time (ISO 8601 format)
        time_zone: Optional time zone for the event
        location: Optional location object with displayName
        attendees: Optional list of attendee info dicts
        categories: Optional list of categories
        show_as: Optional show as status (free, tentative, busy, oof, workingElsewhere)
        user_id: Optional user ID (defaults to 'me')
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Build the update payload
        event_data = {}
        
        if subject is not None:
            event_data["subject"] = subject
        
        if body is not None:
            event_data["body"] = body
        
        # Handle start/end datetime and timezone updates
        if start_datetime is not None or end_datetime is not None or time_zone is not None:
            # Only update start if start_datetime or time_zone is provided
            if start_datetime is not None or time_zone is not None:
                event_data["start"] = {}
                if start_datetime is not None:
                    event_data["start"]["dateTime"] = start_datetime
                if time_zone is not None:
                    event_data["start"]["timeZone"] = time_zone
            
            # Only update end if end_datetime or time_zone is provided
            if end_datetime is not None or time_zone is not None:
                event_data["end"] = {}
                if end_datetime is not None:
                    event_data["end"]["dateTime"] = end_datetime
                if time_zone is not None:
                    event_data["end"]["timeZone"] = time_zone
        
        if location is not None:
            event_data["location"] = location
        
        if attendees is not None:
            event_data["attendees"] = [
                {
                    "emailAddress": attendee.get("emailAddress", {}),
                    "type": attendee.get("type", "required")
                }
                for attendee in attendees
            ]
        
        if categories is not None:
            event_data["categories"] = categories
        
        if show_as is not None:
            event_data["showAs"] = show_as
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/events/{event_id}"
        
        # Make the API call
        result = client.patch(endpoint, json=event_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def get_schedule(
    client,
    Schedules: List[str],
    StartTime: dict,
    EndTime: dict,
    availabilityViewInterval: Optional[str] = None
) -> dict:
    """
    Retrieves free/busy schedule information for specified email addresses within a defined time window.
    
    Args:
        client: The OutlookClient instance
        Schedules: List of email addresses to get schedule for
        StartTime: Start time object {"dateTime": "2026-02-04T09:00:00", "timeZone": "Pacific Standard Time"}
        EndTime: End time object {"dateTime": "2026-02-04T18:00:00", "timeZone": "Pacific Standard Time"}
        availabilityViewInterval: Optional interval in minutes (e.g., "30")
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Build the request payload
        schedule_data = {
            "schedules": Schedules,
            "startTime": StartTime,
            "endTime": EndTime
        }
        
        if availabilityViewInterval is not None:
            schedule_data["availabilityViewInterval"] = int(availabilityViewInterval)
        
        # Make the API call
        endpoint = "/me/calendar/getSchedule"
        result = client.post(endpoint, json=schedule_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }
