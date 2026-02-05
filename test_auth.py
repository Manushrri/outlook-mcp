#!/usr/bin/env python3
"""
Test script to verify OAuth authentication is working.
Run this after setting up your Azure AD app and .env file.
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')

from src.client import get_client


def main():
    print("=" * 60)
    print("Microsoft Outlook MCP - Authentication Test")
    print("=" * 60)
    print()
    
    try:
        # Get client instance
        client = get_client()
        
        # Check if already authenticated
        if client.is_authenticated():
            print("[OK] Already authenticated with cached token")
        else:
            print("No cached token found. Starting authentication...")
            if not client.authenticate_interactive():
                print("[FAILED] Authentication failed!")
                return
        
        # Test API call
        print("\nTesting API connection...")
        user = client.get_me()
        
        print()
        print("[OK] Successfully connected to Microsoft Graph API!")
        print()
        print("User Profile:")
        print(f"  Name: {user.get('displayName', 'N/A')}")
        print(f"  Email: {user.get('mail') or user.get('userPrincipalName', 'N/A')}")
        print(f"  ID: {user.get('id', 'N/A')}")
        print()
        print("=" * 60)
        print("Authentication test PASSED! You're ready to use the MCP server.")
        print("=" * 60)
        
    except ValueError as e:
        print(f"[ERROR] Configuration Error: {e}")
        print()
        print("Please set up your .env file with:")
        print("  OUTLOOK_CLIENT_ID=your_client_id_here")
        print()
        print("See README.md for setup instructions.")
    except Exception as e:
        print(f"[ERROR] {e}")


if __name__ == "__main__":
    main()


