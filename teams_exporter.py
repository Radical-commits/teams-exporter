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
import re
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
import msal
import requests


def strip_html(html_content):
    """Remove HTML tags and extract plain text."""
    if not html_content:
        return ""
    # Remove HTML tags
    text = re.sub(r'<[^>]+>', '', html_content)
    # Decode HTML entities
    text = text.replace('&nbsp;', ' ')
    text = text.replace('&lt;', '<')
    text = text.replace('&gt;', '>')
    text = text.replace('&amp;', '&')
    text = text.replace('&quot;', '"')
    # Clean up whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def format_datetime(iso_datetime):
    """Format ISO datetime to readable format."""
    try:
        dt = datetime.fromisoformat(iso_datetime.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d %H:%M:%S UTC')
    except:
        return iso_datetime


def convert_thread_to_markdown(message, replies):
    """Convert message and replies to markdown format."""
    # Extract message details
    from_data = message.get('from') or {}
    user_data = from_data.get('user') or {}
    author = user_data.get('displayName', 'Unknown')
    created = format_datetime(message.get('createdDateTime', ''))
    subject = message.get('subject') or '(No Subject)'
    body_data = message.get('body') or {}
    body_html = body_data.get('content', '')
    body = strip_html(body_html)

    # Build markdown
    md = []
    md.append(f"# {subject}")
    md.append("")
    md.append(f"**Posted by:** {author}")
    md.append(f"**Date:** {created}")
    md.append("")
    md.append("---")
    md.append("")
    md.append(body)
    md.append("")

    # Add replies (sorted chronologically)
    if replies:
        sorted_replies = sorted(replies, key=lambda r: r.get('createdDateTime', ''))

        md.append("---")
        md.append("")
        md.append(f"## Replies ({len(sorted_replies)})")
        md.append("")

        for idx, reply in enumerate(sorted_replies, 1):
            from_data = reply.get('from') or {}
            user_data = from_data.get('user') or {}
            reply_author = user_data.get('displayName', 'Unknown')
            reply_created = format_datetime(reply.get('createdDateTime', ''))
            reply_body_data = reply.get('body') or {}
            reply_body_html = reply_body_data.get('content', '')
            reply_body = strip_html(reply_body_html)

            md.append(f"### Reply {idx}")
            md.append("")
            md.append(f"**From:** {reply_author}")
            md.append(f"**Date:** {reply_created}")
            md.append("")
            md.append(reply_body)
            md.append("")

    return '\n'.join(md)


def get_access_token_interactive(client_id, tenant_id):
    """
    Acquire access token using interactive browser flow.
    Browser opens automatically for user authentication.

    Args:
        client_id: Azure AD application ID
        tenant_id: Azure AD tenant ID

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

    # Interactive authentication using browser
    print("\nStarting interactive authentication...")
    print("A browser window will open for sign-in.")
    print()

    try:
        result = app.acquire_token_interactive(
            scopes=scopes,
            prompt="select_account"
        )
    except Exception as e:
        print(f"Authentication failed: {e}")
        return None

    if "access_token" in result:
        print("✓ Authentication successful")
        return result["access_token"]

    print(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
    return None


def export_messages(access_token, team_id, channel_id, output_dir="./exports", max_messages=None,
                    fetch_replies=False, max_replies_per_message=None, reply_fetch_delay=0.5):
    """
    Export messages from a channel (user context) to individual thread files.

    Args:
        access_token: Graph API access token (delegated)
        team_id: Team ID
        channel_id: Channel ID
        output_dir: Output directory
        max_messages: Maximum number of messages to export
        fetch_replies: Enable reply fetching (default: False)
        max_replies_per_message: Limit replies per message (optional)
        reply_fetch_delay: Delay between reply requests in seconds (default: 0.5)

    Returns:
        Path to export folder or None if failed
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Create output directory
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Create timestamped subfolder for this export
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    export_folder = os.path.join(output_dir, f"export_{timestamp}")
    Path(export_folder).mkdir(parents=True, exist_ok=True)

    export_start_time = datetime.now().isoformat()
    thread_files = []
    total_replies = 0

    # Fetch messages
    fetch_mode = "with replies" if fetch_replies else "without replies"
    print(f"\nFetching messages ({fetch_mode}){f', limit: {max_messages}' if max_messages else ''}...")

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

            print(f"  Page {page}: {len(batch)} messages (total: {len(messages) + len(batch)})")

            # Process each message and save to individual file
            if fetch_replies:
                for idx, message in enumerate(batch, 1):
                    message_id = message.get('id')

                    # Fetch replies for this message
                    replies = fetch_message_replies(
                        access_token, team_id, channel_id, message_id,
                        max_replies_per_message
                    )

                    # Create thread file (markdown)
                    thread_filename = f"thread_{message_id}.md"
                    thread_filepath = os.path.join(export_folder, thread_filename)

                    # Convert to markdown
                    markdown_content = convert_thread_to_markdown(message, replies)

                    # Save markdown file immediately
                    with open(thread_filepath, 'w', encoding='utf-8') as f:
                        f.write(markdown_content)

                    thread_files.append(thread_filename)
                    total_replies += len(replies)

                    # Progress indicator
                    total_msg_count = len(messages) + idx
                    if replies:
                        print(f"    Thread {total_msg_count}: {len(replies)} replies → {thread_filename}")

                    # Rate limiting delay
                    if reply_fetch_delay > 0:
                        time.sleep(reply_fetch_delay)
            else:
                # When not fetching replies, still save each message to separate file
                for idx, message in enumerate(batch, 1):
                    message_id = message.get('id')
                    thread_filename = f"thread_{message_id}.md"
                    thread_filepath = os.path.join(export_folder, thread_filename)

                    # Convert to markdown (no replies)
                    markdown_content = convert_thread_to_markdown(message, [])

                    # Save markdown file
                    with open(thread_filepath, 'w', encoding='utf-8') as f:
                        f.write(markdown_content)

                    thread_files.append(thread_filename)

            messages.extend(batch)

            # Check limit
            if max_messages and len(messages) >= max_messages:
                # Truncate messages list if over limit
                if len(messages) > max_messages:
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

    # After pagination loop completes, save metadata file
    metadata = {
        "team_id": team_id,
        "channel_id": channel_id,
        "exported_at": export_start_time,
        "message_count": len(messages),
        "reply_count": total_replies,
        "pages_fetched": page,
        "fetch_replies": fetch_replies,
        "auth_type": "delegated (user context)",
        "thread_files": thread_files
    }

    metadata_filepath = os.path.join(export_folder, "_metadata.json")
    with open(metadata_filepath, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)

    return export_folder


def fetch_message_replies(access_token, team_id, channel_id, message_id, max_replies=None):
    """
    Fetch all replies for a specific message.

    Args:
        access_token: Graph API access token (delegated)
        team_id: Team ID
        channel_id: Channel ID
        message_id: Parent message ID
        max_replies: Maximum number of replies to fetch per message (None = unlimited)

    Returns:
        List of reply message objects, or empty list if failed/no replies
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    replies = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies?$top=50"

    while url:
        try:
            response = requests.get(url, headers=headers, timeout=30)

            # 404 means no replies - this is normal
            if response.status_code == 404:
                return []

            # Handle rate limiting
            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 60))
                time.sleep(retry_after)
                continue

            if response.status_code != 200:
                # Log warning but don't crash - return empty list
                return []

            data = response.json()
            batch = data.get("value", [])
            replies.extend(batch)

            # Check limit
            if max_replies and len(replies) >= max_replies:
                replies = replies[:max_replies]
                break

            # Next page
            url = data.get("@odata.nextLink")
            if url:
                time.sleep(0.3)  # Rate limit prevention

        except requests.exceptions.RequestException:
            return []  # Graceful failure

    return replies


def main():
    """Main entry point."""
    load_dotenv()

    # Load configuration
    client_id = os.getenv('CLIENT_ID')
    tenant_id = os.getenv('TENANT_ID')
    team_id = os.getenv('TEAM_ID')
    channel_id = os.getenv('CHANNEL_ID')
    max_messages_str = os.getenv('MAX_MESSAGES')
    output_dir = os.getenv('OUTPUT_DIR', './exports')

    # Reply fetching configuration
    fetch_replies = os.getenv('FETCH_REPLIES', 'false').lower() == 'true'
    max_replies_per_message_str = os.getenv('MAX_REPLIES_PER_MESSAGE')
    max_replies_per_message = int(max_replies_per_message_str) if max_replies_per_message_str else None
    reply_fetch_delay = float(os.getenv('REPLY_FETCH_DELAY', '0.5'))

    # Validate required variables (no CLIENT_SECRET needed for user auth)
    if not all([client_id, tenant_id, team_id, channel_id]):
        print("Error: Missing required environment variables")
        print("Required: CLIENT_ID, TENANT_ID, TEAM_ID, CHANNEL_ID")
        print("\nNote: CLIENT_SECRET is NOT needed for user authentication")
        sys.exit(1)

    max_messages = int(max_messages_str) if max_messages_str else None

    # Authenticate (interactive - user logs in)
    access_token = get_access_token_interactive(client_id, tenant_id)
    if not access_token:
        sys.exit(1)

    # Export (returns folder path, not file path)
    export_folder = export_messages(
        access_token, team_id, channel_id, output_dir, max_messages,
        fetch_replies, max_replies_per_message, reply_fetch_delay
    )

    if export_folder:
        # Read metadata from _metadata.json
        metadata_path = os.path.join(export_folder, "_metadata.json")
        with open(metadata_path, 'r') as f:
            metadata = json.load(f)

        print(f"\n✓ Export complete: {export_folder}")
        print(f"  Messages: {metadata['message_count']}")
        if fetch_replies:
            print(f"  Replies: {metadata['reply_count']}")
        print(f"  Thread files: {len(metadata['thread_files'])}")

        print("\nNote: This export only includes channels you have access to.")
    else:
        print("\n✗ Export failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
