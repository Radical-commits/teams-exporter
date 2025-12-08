#!/usr/bin/env python3
"""
Microsoft Teams Message Exporter (User Context)

Exports messages from Teams channels using delegated permissions.
Only accesses channels the authenticated user can see.
"""

import os
import sys
import json
import time
import threading
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
import msal
import requests


def get_access_token_interactive(client_id, tenant_id, timeout=300):
    """
    Acquire access token using device code flow (interactive).
    User logs in with their account.

    Args:
        client_id: Azure AD application ID
        tenant_id: Azure AD tenant ID
        timeout: Authentication timeout in seconds (default: 300 = 5 minutes)

    Returns:
        Access token string or None if failed
    """
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = [
        "ChannelMessage.Read.All",
        "User.Read"
    ]

    app = msal.PublicClientApplication(
        client_id,
        authority=authority
    )

    # Try to get token from cache first
    accounts = app.get_accounts()
    if accounts:
        print("Found cached account, attempting silent authentication...")
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            print("✓ Authenticated from cache")
            return result["access_token"]

    # Interactive authentication using device code flow
    print("\nStarting interactive authentication...")
    print(f"Timeout: {timeout} seconds ({timeout // 60} minutes)")

    flow = app.initiate_device_flow(scopes=scopes)

    if "user_code" not in flow:
        print(f"Error initiating device flow: {flow.get('error_description', 'Unknown error')}")
        return None

    print(flow["message"])
    print()

    # Use threading to enforce timeout on blocking MSAL call
    result_container = {"result": None, "completed": False}
    start_time = time.time()

    def acquire_token_thread():
        """Thread function to acquire token."""
        try:
            result_container["result"] = app.acquire_token_by_device_flow(flow)
            result_container["completed"] = True
        except Exception as e:
            result_container["result"] = {"error": str(e)}
            result_container["completed"] = True

    # Start authentication in separate thread
    auth_thread = threading.Thread(target=acquire_token_thread, daemon=True)
    auth_thread.start()

    # Wait for completion with timeout
    auth_thread.join(timeout=timeout)
    elapsed = time.time() - start_time

    # Check if thread completed
    if not result_container["completed"]:
        print(f"\n✗ Authentication timeout after {int(elapsed)}s")
        print("Please try again and complete authentication within the time limit.")
        return None

    result = result_container["result"]

    if "access_token" in result:
        print(f"✓ Authentication successful (took {int(elapsed)}s)")
        return result["access_token"]

    print(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
    return None


def export_messages(access_token, team_id, channel_id, output_dir="./exports", max_messages=None):
    """
    Export messages from a channel (user context).

    Args:
        access_token: Graph API access token (delegated)
        team_id: Team ID
        channel_id: Channel ID
        output_dir: Output directory
        max_messages: Maximum number of messages to export

    Returns:
        Path to exported file or None if failed
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Create output directory
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Fetch messages
    print(f"\nFetching messages{f' (limit: {max_messages})' if max_messages else ''}...")

    messages = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages?$top=50"
    page = 0

    while url:
        page += 1

        try:
            response = requests.get(url, headers=headers, timeout=30)

            # Handle rate limiting
            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 60))
                print(f"Rate limited, waiting {retry_after}s...")
                time.sleep(retry_after)
                continue

            if response.status_code != 200:
                print(f"Error {response.status_code}: {response.text}")
                return None

            data = response.json()
            batch = data.get("value", [])
            messages.extend(batch)

            print(f"  Page {page}: {len(batch)} messages (total: {len(messages)})")

            # Check limit
            if max_messages and len(messages) >= max_messages:
                messages = messages[:max_messages]
                print(f"Reached limit of {max_messages} messages")
                break

            # Next page
            url = data.get("@odata.nextLink")
            if url:
                time.sleep(0.5)

        except requests.exceptions.RequestException as e:
            print(f"Network error: {e}")
            return None

    # Save to file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"messages_user_{timestamp}.json"
    filepath = os.path.join(output_dir, filename)

    export_data = {
        "metadata": {
            "team_id": team_id,
            "channel_id": channel_id,
            "exported_at": datetime.now().isoformat(),
            "message_count": len(messages),
            "pages_fetched": page,
            "auth_type": "delegated (user context)"
        },
        "messages": messages
    }

    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(export_data, f, indent=2, ensure_ascii=False)

    return filepath


def main():
    """Main entry point."""
    load_dotenv()

    # Load configuration
    client_id = os.getenv('CLIENT_ID')
    tenant_id = os.getenv('TENANT_ID')
    team_id = os.getenv('TEAM_ID')
    channel_id = os.getenv('CHANNEL_ID')
    max_messages_str = os.getenv('MAX_MESSAGES')
    auth_timeout_str = os.getenv('AUTH_TIMEOUT')
    output_dir = os.getenv('OUTPUT_DIR', './exports')

    # Validate required variables (no CLIENT_SECRET needed for user auth)
    if not all([client_id, tenant_id, team_id, channel_id]):
        print("Error: Missing required environment variables")
        print("Required: CLIENT_ID, TENANT_ID, TEAM_ID, CHANNEL_ID")
        print("\nNote: CLIENT_SECRET is NOT needed for user authentication")
        sys.exit(1)

    max_messages = int(max_messages_str) if max_messages_str else None
    auth_timeout = int(auth_timeout_str) if auth_timeout_str else 300

    # Authenticate (interactive - user logs in)
    access_token = get_access_token_interactive(client_id, tenant_id, auth_timeout)
    if not access_token:
        sys.exit(1)

    # Export
    filepath = export_messages(access_token, team_id, channel_id, output_dir, max_messages)

    if filepath:
        print(f"\n✓ Export complete: {filepath}")
        print("\nNote: This export only includes channels you have access to.")
    else:
        print("\n✗ Export failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
