import requests
import base64
import os
import sys
from config import (
    WEBWORK_USERS_API_URL,
    WEBWORK_API_KEY,
    WEBWORK_API_USER
)

def get_auth_header():
    """Get Basic Authentication header for WebWork API."""
    credentials = f"{WEBWORK_API_USER}:{WEBWORK_API_KEY}"
    encoded_credentials = base64.b64encode(credentials.encode()).decode()
    return {"Authorization": f"Basic {encoded_credentials}"}

def find_users_in_webwork(names_to_find):
    """
    Connects to WebWork, fetches all users, and searches for specific names.

    Args:
        names_to_find (list): A list of names to search for.
    """
    print("Connecting to WebWork to fetch user directory...")
    try:
        response = requests.get(WEBWORK_USERS_API_URL, headers=get_auth_header())
        response.raise_for_status()
        # Handle UTF-8 BOM if present
        response.encoding = 'utf-8-sig'
        all_users = response.json()
        print(f"Successfully downloaded directory with {len(all_users)} users.")
    except requests.exceptions.RequestException as e:
        print(f"Error: Could not fetch user directory from WebWork. {e}", file=sys.stderr)
        return

    print("\n--- Search Results ---")
    names_to_find_lower = [name.lower() for name in names_to_find]
    found_count = 0

    for user in all_users:
        user_fullname = user.get('fullname', '').lower()
        user_email = user.get('email', 'N/A')
        user_id = user.get('id', 'N/A')
        
        for name_to_find_lower in names_to_find_lower:
            # Flexible search: check if the name is contained in the full name
            if name_to_find_lower in user_fullname:
                print(f"\nMatch Found for '{name_to_find_lower.title()}'")
                print(f"  - WebWork Full Name: {user.get('fullname', 'N/A')}")
                print(f"  - Email (ID):        {user_email}")
                print(f"  - WebWork User ID:   {user_id}")
                found_count += 1

    if found_count == 0:
        print("\nNo matches found for the specified names in the WebWork user directory.")
    else:
        print(f"\nFound {found_count} potential match(es).")

def main():
    """Main function to run the user search."""
    # List of names to search for, as provided by the user
    target_names = [
        "Mikaela Gordon",
        "Mike Hammer",
        "Nick Pintozzi",
        "Nicole Nelson"
    ]
    
    print(f"Searching for the following {len(target_names)} users in WebWork:")
    for name in target_names:
        print(f"- {name}")
    print("-" * 20)
    
    find_users_in_webwork(target_names)

if __name__ == "__main__":
    main() 