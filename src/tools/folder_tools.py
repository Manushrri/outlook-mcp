"""
Microsoft Outlook Folder Tools
"""

from typing import Optional


def create_mail_folder(
    client,
    displayName: str,
    isHidden: Optional[bool] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Create a new mail folder.
    Use when you need to organize email into a new folder.
    
    Args:
        client: The OutlookClient instance
        displayName: The display name of the mail folder
        isHidden: Whether the folder is hidden
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
        
        # Build the folder payload
        folder_data = {
            "displayName": displayName
        }
        
        # Add optional fields if provided
        if isHidden is not None:
            folder_data["isHidden"] = isHidden
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/mailFolders"
        
        # Make the API call
        result = client.post(endpoint, json=folder_data)
        
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


def delete_mail_folder(
    client,
    folder_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Delete a mail folder from the user's mailbox.
    Use when you need to remove an existing mail folder.
    
    Args:
        client: The OutlookClient instance
        folder_id: The ID of the mail folder to delete
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
        endpoint = f"/{user}/mailFolders/{folder_id}"
        
        # Make the API call
        result = client.delete(endpoint)
        
        return {
            "successful": True,
            "data": result if result else {"deleted": True}
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }

