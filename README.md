# Teams Message Exporter

Simple Python tool for exporting Microsoft Teams channel messages using the Microsoft Graph API with **user authentication**. Only accesses channels the authenticated user has permission to see.

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
TENANT_ID=your_tenant_id
TEAM_ID=your_team_id
CHANNEL_ID=your_channel_id
```

**Note:** CLIENT_SECRET is NOT needed for user authentication.

Optional:
```bash
MAX_MESSAGES=100        # Limit number of messages (recommended for speed)
AUTH_TIMEOUT=300        # Authentication timeout in seconds (default: 300 = 5 minutes)
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
2. Create new registration (or use existing)
3. Add **Delegated permissions** (not Application):
   - `ChannelMessage.Read.All`
   - `User.Read`
4. Grant admin consent for these permissions
5. **Enable public client flows**:
   - Go to Authentication tab
   - Scroll to "Advanced settings"
   - Set "Allow public client flows" to **YES**
6. Add platform (if not already added):
   - Authentication → Add a platform → Mobile and desktop applications
   - Check: `https://login.microsoftonline.com/common/oauth2/nativeclient`
7. Copy CLIENT_ID and TENANT_ID to `.env`

**Important:** No client secret needed for user authentication.

## Usage

```bash
python teams_exporter.py
```

**First run:** You'll be prompted to authenticate interactively:
```
Starting interactive authentication...
Timeout: 300 seconds (5 minutes)
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXX to authenticate.
```

1. Open the URL in your browser
2. Enter the code shown
3. Sign in with your Microsoft account
4. Authorize the application
5. Complete within the timeout period (default: 5 minutes)

**Subsequent runs:** The token is cached, so you won't need to authenticate again unless it expires.

**Timeout:** If authentication takes too long, the script will exit with a timeout error. You can adjust the timeout using the `AUTH_TIMEOUT` environment variable.

Output:
```
✓ Authenticated from cache
Fetching messages (limit: 100)...
  Page 1: 50 messages (total: 50)
  Page 2: 50 messages (total: 100)
Reached limit of 100 messages

✓ Export complete: ./exports/messages_user_20251205_170147.json

Note: This export only includes channels you have access to.
```

## Export Format

```json
{
  "metadata": {
    "team_id": "...",
    "channel_id": "...",
    "exported_at": "2025-12-05T17:01:47.123456",
    "message_count": 100,
    "pages_fetched": 2,
    "auth_type": "delegated (user context)"
  },
  "messages": [...]
}
```

## Performance

- With `MAX_MESSAGES=100`: ~5 seconds, 2 pages
- Without limit: Fetches all messages in channel (may take minutes for large channels)

## Security Note

This tool uses **delegated permissions** and **user authentication**:
- Only accesses channels the authenticated user can see
- Requires interactive sign-in (device code flow)
- Does NOT use application permissions or client secrets
- Access is limited by the user's Teams permissions
