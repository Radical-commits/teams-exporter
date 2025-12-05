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

First run prompts for interactive authentication via device code flow. Token is cached for subsequent runs.

## Architecture

### Core Functions

**`get_access_token_interactive(client_id, tenant_id)`** - teams_exporter.py:20
- Authenticates using MSAL device code flow (interactive)
- Tries token cache first for silent authentication
- Prompts user to visit microsoft.com/devicelogin and enter code
- Returns access token or None

**`export_messages(access_token, team_id, channel_id, output_dir, max_messages)`** - teams_exporter.py:68
- Fetches messages from Graph API with pagination
- Handles rate limiting (429 responses)
- Stops early if max_messages limit reached
- Saves to timestamped JSON file with auth_type metadata

**`main()`** - teams_exporter.py:158
- Loads environment variables
- Validates required configuration (no CLIENT_SECRET required)
- Orchestrates interactive authentication and export

### Microsoft Graph API

**Endpoint**: `GET /teams/{team-id}/channels/{channel-id}/messages`

**Pagination**: Uses `$top=50` and follows `@odata.nextLink`

**Rate Limiting**: Auto-retries with `Retry-After` header value

### Required Permissions

Delegated permissions (with admin consent):
- `Channel.ReadBasic.All` (Delegated)
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
- MAX_MESSAGES=100: ~5 seconds, 2 API pages
- No limit: Full channel scan (minutes for large channels)

Messages are returned newest-first, so limiting gives most recent content.

## File Structure

- `teams_exporter.py`: Main script with user authentication (~197 lines)
- `create_threads.py`: Convert exported JSON to threaded Markdown format
- `.env`: Configuration (gitignored)
- `requirements.txt`: Dependencies (msal, requests, python-dotenv)
- `exports/`: Output directory for JSON files

## Dependencies

- **msal**: Microsoft Authentication Library
- **requests**: HTTP client for Graph API
- **python-dotenv**: Environment variable management
