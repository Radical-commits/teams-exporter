# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Simple Python tool for exporting Microsoft Teams channel messages via Microsoft Graph API. Uses **device code flow** (interactive user authentication) with **delegated permissions** - only accesses channels the authenticated user can see.

## Development Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Running the Exporter

```bash
python teams_exporter.py
```

Requires `.env` file with:
- Azure AD credentials (CLIENT_ID, TENANT_ID) - **No CLIENT_SECRET needed**
- Teams identifiers (TEAM_ID, CHANNEL_ID)
- Optional: MAX_MESSAGES (limits export, recommended for speed)
- Optional: FETCH_REPLIES (enable reply fetching, default: false)
- Optional: MAX_REPLIES_PER_MESSAGE, REPLY_FETCH_DELAY

First run opens browser automatically for authentication. Token is cached for subsequent runs.

Export creates a timestamped folder with individual thread files and a metadata summary.

## Architecture

### Core Functions

**`strip_html(html_content)`** - teams_exporter.py:21
- Removes HTML tags and decodes HTML entities
- Used by markdown conversion to extract plain text from message bodies

**`format_datetime(iso_datetime)`** - teams_exporter.py:38
- Formats ISO datetime strings to readable format (YYYY-MM-DD HH:MM:SS UTC)

**`convert_thread_to_markdown(message, replies)`** - teams_exporter.py:47
- Converts message and replies to markdown format
- Handles None values gracefully
- Sorts replies chronologically (oldest first)
- Returns formatted markdown string

**`get_access_token_interactive(client_id, tenant_id)`** - teams_exporter.py:94
- Authenticates using MSAL interactive browser flow
- Tries token cache first for silent authentication
- Opens browser automatically to Microsoft login page
- Handles OAuth redirect internally (no manual code entry)
- Returns access token or None if failed

**`export_messages(access_token, team_id, channel_id, output_dir, max_messages, fetch_replies, max_replies_per_message, reply_fetch_delay)`** - teams_exporter.py:148
- Creates timestamped subfolder for export
- Fetches messages from Graph API with pagination
- Handles rate limiting (429 responses)
- For each message: fetches replies (if enabled), converts to markdown, and saves to individual .md file
- Creates _metadata.json with export summary
- Returns export folder path (not single file path)

**`fetch_message_replies(access_token, team_id, channel_id, message_id, max_replies)`** - teams_exporter.py:308
- Fetches all replies for a specific message ID
- Handles pagination with @odata.nextLink
- Handles rate limiting (429 responses)
- Returns empty list on errors (graceful degradation)
- Returns list of reply message objects

**`main()`** - teams_exporter.py:368
- Loads environment variables (including reply fetching config)
- Validates required configuration (no CLIENT_SECRET required)
- Orchestrates interactive authentication and export
- Displays export statistics from metadata file

### Microsoft Graph API

**Main Endpoint**: `GET /teams/{team-id}/channels/{channel-id}/messages`
- Fetches top-level messages

**Replies Endpoint**: `GET /teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies`
- Fetches replies for a specific message (when FETCH_REPLIES=true)

**Pagination**: Uses `$top=50` and follows `@odata.nextLink`

**Rate Limiting**: Auto-retries with `Retry-After` header value

### Required Permissions

Delegated permissions (with admin consent):
- `ChannelMessage.Read.All` (Delegated)
- `User.Read` (Delegated)

### Azure AD App Configuration

Required settings:
- **Authentication → Allow public client flows**: YES
- **Platform**: Mobile and desktop applications
- **Redirect URI**: `https://login.microsoftonline.com/common/oauth2/nativeclient`
- **No client secret needed**

### Security Model

- Uses **user context authentication** (delegated permissions)
- Only accesses channels the authenticated user has permission to see
- No application-level access to all Teams data
- Requires interactive sign-in

## Performance Optimization

The `MAX_MESSAGES` parameter stops pagination early:
- Without replies: MAX_MESSAGES=100 in ~5 seconds (2 API pages)
- With replies: MAX_MESSAGES=100 in ~1-3 minutes (100+ API calls for replies)
- No limit: Full channel scan (minutes to hours for large channels)

Messages are returned newest-first, so limiting gives most recent content.

**Reply fetching impact:**
- Opt-in by default (FETCH_REPLIES=false)
- Each message requires additional API call for replies
- Use REPLY_FETCH_DELAY (default 0.5s) to prevent rate limiting

## File Structure

- `teams_exporter.py`: Main script with user authentication and markdown export (~425 lines)
- `thread_to_markdown.py`: Standalone converter (functionality now integrated into main script)
- `.env`: Configuration (gitignored)
- `requirements.txt`: Dependencies (msal, requests, python-dotenv)
- `exports/`: Output directory
  - `export_YYYYMMDD_HHMMSS/`: Timestamped export folder
    - `_metadata.json`: Export summary (team, channel, counts, file list)
    - `thread_*.md`: Individual markdown thread files (message + replies, chronologically sorted)

## Dependencies

- **msal**: Microsoft Authentication Library
- **requests**: HTTP client for Graph API
- **python-dotenv**: Environment variable management
