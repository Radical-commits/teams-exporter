# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Simple Python tool for exporting Microsoft Teams channel messages via Microsoft Graph API. Uses client credentials flow (app-only authentication).

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
- Azure AD credentials (CLIENT_ID, CLIENT_SECRET, TENANT_ID)
- Teams identifiers (TEAM_ID, CHANNEL_ID)
- Optional: MAX_MESSAGES (limits export, recommended for speed)

## Architecture

### Core Functions

**`get_access_token(client_id, client_secret, tenant_id)`** - teams_exporter.py:20
- Authenticates using MSAL client credentials flow
- Returns access token or None

**`export_messages(access_token, team_id, channel_id, output_dir, max_messages)`** - teams_exporter.py:38
- Fetches messages from Graph API with pagination
- Handles rate limiting (429 responses)
- Stops early if max_messages limit reached
- Saves to timestamped JSON file

**`main()`** - teams_exporter.py:127
- Loads environment variables
- Validates required configuration
- Orchestrates authentication and export

### Microsoft Graph API

**Endpoint**: `GET /teams/{team-id}/channels/{channel-id}/messages`

**Pagination**: Uses `$top=50` and follows `@odata.nextLink`

**Rate Limiting**: Auto-retries with `Retry-After` header value

### Required Permissions

Application permissions (with admin consent):
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`

## Performance Optimization

The `MAX_MESSAGES` parameter stops pagination early:
- MAX_MESSAGES=100: ~5 seconds, 2 API pages
- No limit: Full channel scan (minutes for large channels)

Messages are returned newest-first, so limiting gives most recent content.

## File Structure

- `teams_exporter.py`: Main script (166 lines)
- `.env`: Configuration (gitignored)
- `requirements.txt`: Dependencies (msal, requests, python-dotenv)
- `exports/`: Output directory for JSON files

## Dependencies

- **msal**: Microsoft Authentication Library
- **requests**: HTTP client for Graph API
- **python-dotenv**: Environment variable management
