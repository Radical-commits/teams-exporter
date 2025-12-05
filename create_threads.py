#!/usr/bin/env python3
"""
Convert Teams messages JSON export to threaded Markdown discussions
"""

import json
import sys
from datetime import datetime
from pathlib import Path


def parse_html_content(html_text):
    """Extract plain text from HTML content."""
    if not html_text:
        return ""

    # Simple HTML tag removal (basic approach)
    import re
    text = re.sub(r'<br\s*/?>', '\n', html_text)
    text = re.sub(r'<[^>]+>', '', text)
    text = re.sub(r'&nbsp;', ' ', text)
    text = re.sub(r'&lt;', '<', text)
    text = re.sub(r'&gt;', '>', text)
    text = re.sub(r'&amp;', '&', text)
    return text.strip()


def format_message(msg, indent=0):
    """Format a single message."""
    indent_str = "  " * indent

    # Extract message details
    msg_id = msg.get('id', 'unknown')
    created = msg.get('createdDateTime', '')

    # Safely get author name
    from_field = msg.get('from')
    if from_field and isinstance(from_field, dict):
        user = from_field.get('user', {})
        if user:
            author = user.get('displayName', 'Unknown')
        else:
            author = 'System'
    else:
        author = 'Unknown'
    subject = msg.get('subject') or ''

    # Parse timestamp
    if created:
        dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
        timestamp = dt.strftime('%Y-%m-%d %H:%M')
    else:
        timestamp = 'Unknown time'

    # Get message content
    body = msg.get('body', {})
    content_type = body.get('contentType', 'text')
    content = body.get('content', '')

    if content_type == 'html':
        content = parse_html_content(content)

    # Format the message
    lines = []
    if subject and indent == 0:  # Only show subject for root messages
        lines.append(f"{indent_str}### {subject}")
        lines.append("")

    lines.append(f"{indent_str}**{author}** • {timestamp}")
    lines.append("")

    if content:
        for line in content.split('\n'):
            if line.strip():
                lines.append(f"{indent_str}{line}")
        lines.append("")

    return '\n'.join(lines)


def build_thread_tree(messages):
    """Build a tree structure of messages based on replyToId."""
    # Create lookup dictionaries
    msg_by_id = {msg['id']: msg for msg in messages}
    root_messages = []
    replies_by_parent = {}

    for msg in messages:
        reply_to = msg.get('replyToId')
        if reply_to:
            if reply_to not in replies_by_parent:
                replies_by_parent[reply_to] = []
            replies_by_parent[reply_to].append(msg)
        else:
            root_messages.append(msg)

    return root_messages, replies_by_parent, msg_by_id


def format_thread(msg, replies_by_parent, indent=0):
    """Recursively format a message and its replies."""
    lines = [format_message(msg, indent)]

    msg_id = msg['id']
    if msg_id in replies_by_parent:
        replies = sorted(replies_by_parent[msg_id],
                        key=lambda x: x.get('createdDateTime', ''))

        for reply in replies:
            lines.append(format_thread(reply, replies_by_parent, indent + 1))

    return '\n'.join(lines)


def main():
    if len(sys.argv) < 2:
        print("Usage: python create_threads.py <json_file>")
        sys.exit(1)

    input_file = sys.argv[1]

    # Read JSON
    print(f"Reading {input_file}...")
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    messages = data.get('messages', [])
    metadata = data.get('metadata', {})

    print(f"Found {len(messages)} messages")

    # Build thread tree
    root_messages, replies_by_parent, msg_by_id = build_thread_tree(messages)

    print(f"Found {len(root_messages)} root discussion threads")

    # Sort root messages by date (newest first)
    root_messages = sorted(root_messages,
                          key=lambda x: x.get('createdDateTime', ''),
                          reverse=True)

    # Generate Markdown
    output_lines = []
    output_lines.append("# Teams Channel Discussions")
    output_lines.append("")
    output_lines.append(f"**Exported**: {metadata.get('exported_at', 'Unknown')}")
    output_lines.append(f"**Total Messages**: {metadata.get('message_count', 0)}")
    output_lines.append("")
    output_lines.append("---")
    output_lines.append("")

    for i, root_msg in enumerate(root_messages, 1):
        output_lines.append(f"## Thread {i}")
        output_lines.append("")
        output_lines.append(format_thread(root_msg, replies_by_parent))
        output_lines.append("---")
        output_lines.append("")

    # Write output
    output_file = Path(input_file).parent / f"{Path(input_file).stem}_threads.md"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_lines))

    print(f"\n✓ Threads created: {output_file}")
    print(f"  Root threads: {len(root_messages)}")
    print(f"  Total messages: {len(messages)}")


if __name__ == "__main__":
    main()
