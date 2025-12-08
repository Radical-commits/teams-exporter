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
OUTPUT_DIR=./exports    # Output directory

# Reply Fetching (opt-in, significantly slower)
FETCH_REPLIES=false              # Enable fetching replies for each message (default: false)
MAX_REPLIES_PER_MESSAGE=         # Optional: Limit replies per message
REPLY_FETCH_DELAY=0.5            # Delay between reply requests (rate limiting, default: 0.5s)
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

**First run:** A browser window will open automatically for authentication:
```
Starting interactive authentication...
A browser window will open for sign-in.
```

1. Browser opens automatically to Microsoft login page
2. Sign in with your Microsoft account
3. Authorize the application
4. Browser shows success message
5. Return to terminal - export begins automatically

**Subsequent runs:** The token is cached, so you won't need to authenticate again unless it expires.

Output:
```
✓ Authenticated from cache
Fetching messages (without replies, limit: 100)...
  Page 1: 50 messages (total: 50)
  Page 2: 50 messages (total: 100)
Reached limit of 100 messages

✓ Export complete: ./exports/export_20251208_184118
  Messages: 100
  Thread files: 100

Note: This export only includes channels you have access to.
```

## Export Format

Each export creates a timestamped folder with individual markdown thread files:

```
./exports/
  export_20251208_184118/        # Timestamped export folder
    _metadata.json               # Export summary
    thread_1765191244192.md      # Individual markdown thread files (message + replies)
    thread_1764894518606.md
    ...
```

**Metadata file (`_metadata.json`):**
```json
{
  "team_id": "...",
  "channel_id": "...",
  "exported_at": "2025-12-08T18:41:18.141011",
  "message_count": 100,
  "reply_count": 0,
  "pages_fetched": 2,
  "fetch_replies": false,
  "auth_type": "delegated (user context)",
  "thread_files": [
    "thread_1765191244192.md",
    "thread_1764894518606.md",
    ...
  ]
}
```

**Thread file format (Markdown):**
```markdown
# Flow- Caller number attribute

**Posted by:** John Doe
**Date:** 2025-12-08 10:54:04 UTC

---

This is the main message body with HTML tags removed.

---

## Replies (2)

### Reply 1

**From:** Jane Smith
**Date:** 2025-12-08 11:15:00 UTC

This is the first reply.

### Reply 2

**From:** Bob Johnson
**Date:** 2025-12-08 12:30:00 UTC

This is the second reply.
```

## Reply Fetching (Optional)

By default, the exporter only fetches top-level messages without replies. To fetch complete conversation threads with replies:

```bash
FETCH_REPLIES=true
```

**Performance impact:**
- **Without replies:** 100 messages in ~5 seconds
- **With replies:** 100 messages in ~1-3 minutes (depends on reply count)
- Each message requires an additional API call to fetch replies

**Optional configuration:**
```bash
MAX_REPLIES_PER_MESSAGE=50   # Limit replies per message (prevents runaway fetches)
REPLY_FETCH_DELAY=0.5        # Delay between requests in seconds (rate limiting prevention)
```

**Output with replies:**
```
✓ Authenticated from cache
Fetching messages (with replies, limit: 100)...
  Page 1: 50 messages (total: 50)
    Thread 1: 3 replies → thread_1765191244192.md
    Thread 2: 0 replies → thread_1765191244193.md
    Thread 5: 15 replies → thread_1765191244250.md
    ...
  Page 2: 50 messages (total: 100)
    ...
Reached limit of 100 messages

✓ Export complete: ./exports/export_20251208_194523
  Messages: 100
  Replies: 247
  Thread files: 100
```

## Performance

- **Without replies:** 100 messages in ~5 seconds, 2 API pages
- **With replies:** 100 messages in ~1-3 minutes (100+ additional API calls)
- **Without limit:** Fetches all messages in channel (may take minutes to hours for large channels)

## Security Note

This tool uses **delegated permissions** and **user authentication**:
- Only accesses channels the authenticated user can see
- Requires interactive sign-in (device code flow)
- Does NOT use application permissions or client secrets
- Access is limited by the user's Teams permissions
