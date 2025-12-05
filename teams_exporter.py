#!/usr/bin/env python3
"""
Microsoft Teams Message Exporter

Exports messages from a Teams channel using Microsoft Graph API.
Requires: TEAM_ID, CHANNEL_ID, and Azure AD credentials in .env
"""

import os
import sys
import json
import time
from datetime import datetime, timezone
from pathlib import Path
from dotenv import load_dotenv
import msal
import requests


def get_access_token(client_id, client_secret, tenant_id):
    """Acquire access token using client credentials flow."""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        return result["access_token"]

    print(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
    return None


def export_messages(access_token, team_id, channel_id, output_dir="./exports", max_messages=None):
    """
    Export messages from a channel.

    Args:
        access_token: Graph API access token
        team_id: Team ID
        channel_id: Channel ID
        output_dir: Output directory
        max_messages: Maximum number of messages to export (None = all)

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
    print(f"Fetching messages{f' (limit: {max_messages})' if max_messages else ''}...")

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
                time.sleep(0.5)  # Rate limit protection

        except requests.exceptions.RequestException as e:
            print(f"Network error: {e}")
            return None

    # Save to file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"messages_{timestamp}.json"
    filepath = os.path.join(output_dir, filename)

    export_data = {
        "metadata": {
            "team_id": team_id,
            "channel_id": channel_id,
            "exported_at": datetime.now().isoformat(),
            "message_count": len(messages),
            "pages_fetched": page
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
    client_secret = os.getenv('CLIENT_SECRET')
    tenant_id = os.getenv('TENANT_ID')
    team_id = os.getenv('TEAM_ID')
    channel_id = os.getenv('CHANNEL_ID')
    max_messages_str = os.getenv('MAX_MESSAGES')
    output_dir = os.getenv('OUTPUT_DIR', './exports')

    # Validate required variables
    if not all([client_id, client_secret, tenant_id, team_id, channel_id]):
        print("Error: Missing required environment variables")
        print("Required: CLIENT_ID, CLIENT_SECRET, TENANT_ID, TEAM_ID, CHANNEL_ID")
        sys.exit(1)

    max_messages = int(max_messages_str) if max_messages_str else None

    # Authenticate
    print("Authenticating...")
    access_token = get_access_token(client_id, client_secret, tenant_id)
    if not access_token:
        sys.exit(1)

    # Export
    filepath = export_messages(access_token, team_id, channel_id, output_dir, max_messages)

    if filepath:
        print(f"\n✓ Export complete: {filepath}")
    else:
        print("\n✗ Export failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
