# Teams Message Exporter

Simple Python tool for exporting Microsoft Teams channel messages using the Microsoft Graph API.

## Setup

1. Create virtual environment and install dependencies:
```bash
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

2. Configure credentials:
```bash
cp .env.example .env
# Edit .env with your credentials
```

## Configuration

Required variables in `.env`:
```bash
CLIENT_ID=your_azure_ad_app_id
CLIENT_SECRET=your_app_secret
TENANT_ID=your_tenant_id
TEAM_ID=your_team_id
CHANNEL_ID=your_channel_id
```

Optional:
```bash
MAX_MESSAGES=100        # Limit number of messages (recommended for speed)
OUTPUT_DIR=./exports    # Output directory
```

## Finding Team and Channel IDs

### From Teams Web App (easiest)
1. Open Teams in browser: https://teams.microsoft.com
2. Navigate to your channel
3. Copy IDs from URL:
   ```
   https://teams.microsoft.com/l/channel/19%3A...%40thread.tacv2/...?groupId=abc-123...
   ```
   - `groupId=abc-123...` → TEAM_ID
   - `19%3A...%40thread.tacv2` (URL decoded: `19:...@thread.tacv2`) → CHANNEL_ID

## Azure AD Setup

1. Azure Portal → Azure Active Directory → App registrations
2. Create new registration
3. Add **Application permissions** (not Delegated):
   - `Channel.ReadBasic.All`
   - `ChannelMessage.Read.All`
4. Grant admin consent
5. Create client secret under "Certificates & secrets"
6. Copy IDs to `.env`

## Usage

```bash
python teams_exporter.py
```

Output:
```
Authenticating...
Fetching messages (limit: 100)...
  Page 1: 50 messages (total: 50)
  Page 2: 50 messages (total: 100)
Reached limit of 100 messages

✓ Export complete: ./exports/messages_20251205_170147.json
```

## Export Format

```json
{
  "metadata": {
    "team_id": "...",
    "channel_id": "...",
    "exported_at": "2025-12-05T17:01:47.123456",
    "message_count": 100,
    "pages_fetched": 2
  },
  "messages": [...]
}
```

## Performance

- With `MAX_MESSAGES=100`: ~5 seconds, 2 pages
- Without limit: Fetches all messages in channel (may take minutes for large channels)
