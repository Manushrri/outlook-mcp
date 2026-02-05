"""
Microsoft Outlook Profile Tools
"""

from typing import Optional


def get_profile(
    client,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves the Microsoft Outlook profile for a specified user.
    
    Args:
        client: The OutlookClient instance
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
        endpoint = f"/{user}"
        
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

