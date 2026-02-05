"""
Microsoft Outlook Attachment Tools
"""

import base64
from typing import Optional


def download_outlook_attachment(
    client,
    message_id: str,
    attachment_id: str,
    file_name: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Downloads a specific file attachment from an email message in a Microsoft Outlook mailbox.
    The attachment must contain 'contentBytes' (binary data) and not be a link or embedded item.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message containing the attachment
        attachment_id: The ID of the attachment to download
        file_name: The name to save the file as
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
        endpoint = f"/{user}/messages/{message_id}/attachments/{attachment_id}"
        
        # Get the attachment metadata and content
        result = client.get(endpoint)
        
        # Check if it has contentBytes
        if "contentBytes" not in result:
            return {
                "successful": False,
                "data": {},
                "error": "Attachment does not contain downloadable content (contentBytes). It may be a link or embedded item."
            }
        
        # Decode and save the file
        content_bytes = base64.b64decode(result["contentBytes"])
        
        with open(file_name, "wb") as f:
            f.write(content_bytes)
        
        return {
            "successful": True,
            "data": {
                "file_name": file_name,
                "size": len(content_bytes),
                "content_type": result.get("contentType", "unknown"),
                "name": result.get("name", file_name)
            }
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }

