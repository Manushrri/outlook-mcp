"""
Microsoft Outlook Email Rule Tools
"""

from typing import Optional


def create_email_rule(
    client,
    displayName: str,
    conditions: dict,
    actions: dict,
    isEnabled: Optional[bool] = None,
    sequence: Optional[int] = None
) -> dict:
    """
    Create email rule filter with conditions and actions.
    
    Args:
        client: The OutlookClient instance
        displayName: The display name of the rule
        conditions: The conditions that trigger the rule
        actions: The actions to perform when conditions are met
        isEnabled: Whether the rule is enabled (default True)
        sequence: The order of the rule in the rule list
    
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
        
        # Build the rule payload
        rule_data = {
            "displayName": displayName,
            "conditions": conditions,
            "actions": actions
        }
        
        # Add optional fields if provided
        if isEnabled is not None:
            rule_data["isEnabled"] = isEnabled
        if sequence is not None:
            rule_data["sequence"] = sequence
        
        # Endpoint for inbox rules
        endpoint = "/me/mailFolders/inbox/messageRules"
        
        # Make the API call
        result = client.post(endpoint, json=rule_data)
        
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

