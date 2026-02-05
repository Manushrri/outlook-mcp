# Microsoft Outlook MCP Server

A Model Context Protocol (MCP) server for Microsoft Outlook integration using Microsoft Graph API.

## Setup

### 1. Create an Azure AD Application (App Registration)

To let this MCP server call Microsoft Graph on behalf of a user, you must register an app in Azure and get its **Client ID** (and optionally a **Client Secret**).

1. Go to the Azure Portal (`https://portal.azure.com`).
2. Navigate to **Azure Active Directory** → **App registrations**.
3. Click **New registration**.
4. Configure your app:
   - **Name**: `Outlook MCP Server` (or any name you like).
   - **Supported account types** (choose one):
     - **Accounts in any organizational directory and personal Microsoft accounts** – works for both work/school and personal Outlook accounts.
     - **Personal Microsoft accounts only** – if you only care about Outlook.com / personal accounts.
   - You do **not** need to configure a redirect URI for device code flow, but it’s safe to also add a public client redirect:
     - Platform: **Mobile and desktop applications**
     - Redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Click **Register**.

After registration:
- You will use the **Application (client) ID** as `OUTLOOK_CLIENT_ID`.
- The app will request scopes at runtime; the **user sees a consent screen once** during first login and does not need to pick scopes each time.

### 2. Configure Microsoft Graph API Permissions (Scopes)

1. In your app registration, go to **API permissions**.
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
3. Add these permissions:
   - `User.Read` – sign in and read user profile.
   - `Mail.Read` – read user mail.
   - `Mail.ReadWrite` – read and write user mail.
   - `Mail.Send` – send mail as the user.
   - `Calendars.Read` – read user calendars.
   - `Calendars.ReadWrite` – read and write user calendars.
   - `Contacts.Read` – read user contacts.
   - `Contacts.ReadWrite` – read and write user contacts.
   - `MailboxSettings.Read` – read mailbox settings (time zone, automatic replies, etc.).
   - `MailboxSettings.ReadWrite` – read and write mailbox settings.
   - `offline_access` – issue refresh tokens so the MCP can silently refresh access tokens.
4. Click **Grant admin consent** (if you are an admin).  
   - If you cannot grant admin consent, each user will see the consent screen the first time they sign in and accept the permissions themselves.

These scopes are hard‑coded in `src/config.py` and requested automatically by the MCP; users do **not** have to manually choose scopes after the first consent.

### 3. Get Your Client ID and (Optional) Client Secret

1. In your app registration, go to **Overview**.
2. Copy **Application (client) ID** → this is required.
3. (Optional) If you also want to support confidential client flows, create a client secret:
   - Go to **Certificates & secrets**.
   - Click **New client secret**.
   - Add a description and choose expiration.
   - Copy the **Value** (you will not see it again).
4. Add these to your `.env` file in the project root:

```env
OUTLOOK_CLIENT_ID=your-client-id-here
OUTLOOK_CLIENT_SECRET=your-optional-secret-here
OUTLOOK_REDIRECT_URI=https://login.microsoftonline.com/common/oauth2/nativeclient
```

If `OUTLOOK_CLIENT_SECRET` is not set, the MCP uses the **public client + device code flow**, which is what this project is optimized for.

### 4. Install Dependencies

```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 5. First-Time Authentication (Device Code Flow)

Run the server to authenticate the first user:

```bash
python run_server.py
```

On first run:
1. The server starts and triggers the **device code flow**.
2. The terminal will show a **code** and a URL (for example, `https://microsoft.com/devicelogin`).
3. Open the URL in a browser, sign in with your Microsoft account, and enter the code.
4. Microsoft shows a **permission/consent screen** for the scopes listed above. Accept them once.
5. The MCP receives an **access token + refresh token** and saves them to `.token_cache.json`.

After this:
- Future calls use the cached token silently.
- When the access token expires, MSAL uses the refresh token automatically (no extra login needed).
- You only need to re‑authenticate if the cache is deleted or the refresh token is no longer valid.

## Configuration for Cursor/Claude

Add to your MCP settings:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "python",
      "args": ["run_server.py"],
      "cwd": "C:\\Users\\manus\\OneDrive\\Desktop\\microsoft oulook"
    }
  }
}
```

## Project Structure

```
microsoft-outlook/
├── README.md              # Project documentation and setup
├── requirements.txt       # Python dependencies
├── mcp-config.json        # MCP configuration for host tools (Cursor/Claude)
├── tools_manifest.json    # Definition of all MCP tools (IDs, schemas, descriptions)
├── run_server.py          # MCP server entry point
├── test_auth.py           # Helper script for testing authentication (optional)
├── test_attachment.txt    # Sample file used for attachment testing
├── src/
│   ├── __init__.py
│   ├── client.py          # Outlook OAuth/MSAL client (device code flow + token cache)
│   ├── config.py          # Central configuration (client ID, authority, scopes, etc.)
│   ├── main.py            # FastMCP server wiring and tool registration
│   └── tools/
│       ├── __init__.py          # Aggregates and exports all tool functions
│       ├── mail_tools.py        # Email send/reply/draft/search/move/update helpers
│       ├── list_tools.py        # Generic list tools (messages, events, contacts, folders, attachments, reminders)
│       ├── calendar_tools.py    # Calendar and event operations (create/update/delete, schedule, attachments)
│       ├── contact_tools.py     # Contact and contact-folder operations
│       ├── folder_tools.py      # Mail folder create/delete helpers
│       ├── attachment_tools.py  # Download attachment helper
│       ├── category_tools.py    # Category (master category list) tools
│       ├── profile_tools.py     # User profile helper
│       ├── rule_tools.py        # Mail rule creation
│       └── settings_tools.py    # Mailbox settings, delta, mail tips, languages, time zones
└── test/                       # (Optional) test-related artifacts
```

## Available Scopes

| Scope | Description |
|-------|-------------|
| `User.Read` | Read user profile |
| `Mail.Read` | Read emails |
| `Mail.ReadWrite` | Read/write emails |
| `Mail.Send` | Send emails |
| `Calendars.Read` | Read calendar events |
| `Calendars.ReadWrite` | Read/write calendar events |
| `Contacts.Read` | Read contacts |
| `Contacts.ReadWrite` | Read/write contacts |
| `MailboxSettings.Read` | Read mailbox settings |
| `MailboxSettings.ReadWrite` | Read/write mailbox settings |
| `offline_access` | Allow issuing refresh tokens for silent token refresh |

## Available Tools (43 total)

All tools are registered in `tools_manifest.json` and exposed by the MCP server. Below is a high-level overview, grouped by area.

### Mail tools

- **add_mail_attachment**: Attach a small file (<3 MB) to an existing message using `message_id`.
- **create_draft**: Create a draft email (subject, body, recipients, optional attachment).
- **create_draft_reply**: Create a draft reply to an existing message.
- **get_message**: Get one email message by `message_id` (optionally with headers).
- **list_messages**: List messages from a folder (`inbox`, `drafts`, `sentitems`, etc.) with rich filters.
- **search_messages**: Search messages by text, sender, subject, attachments, and pagination.
- **move_message**: Move a message to another folder by `message_id` and destination folder id.
- **reply_email**: Send a plain text reply to a received message.
- **send_email**: Send a new email (subject/body/recipients, optional attachment).
- **update_email**: Update a **draft** message (subject, body, recipients, importance).

### Attachment tools

- **download_attachment**: Download a specific attachment by `message_id` + `attachment_id` to a local file.
- **list_attachments**: List attachment metadata (name, size, type) for a message.
- **add_event_attachment**: Attach a file or item to a calendar event by `event_id`.
- **list_event_attachments**: List attachments for a specific calendar event.

### Calendar tools

- **create_calendar**: Create a new calendar.
- **list_calendars**: List calendars for the signed-in user.
- **create_event**: Create a new event (subject, body, start/end, time zone, attendees, etc.).
- **list_events**: List events with filters, ordering, selection, and timezone preference.
- **get_event**: Get full details for a specific event by `event_id`.
- **delete_event**: Delete an event (optionally notify attendees).
- **update_calendar_event**: Update subject, body, time, location, attendees, categories, or show-as.
- **list_reminders**: Get reminders for events in a time range.
- **get_schedule**: Get free/busy availability for email addresses over a time window.

### Contact tools

- **create_contact**: Create a new contact in the default contacts folder.
- **create_contact_folder**: Create a new contact folder.
- **list_contacts**: List contacts from the default or a specific contact folder.
- **get_contact**: Get details for a contact by `contact_id`.
- **delete_contact**: Delete a contact by `contact_id`.
- **update_contact**: Update an existing contact (name, emails, phones, company info, etc.).
- **get_contact_folders**: List contact folders with optional filters and expansions.

### Folder and rule tools

- **create_mail_folder**: Create a new mail folder.
- **delete_mail_folder**: Delete an existing mail folder by `folder_id`.
- **list_mail_folders**: List top-level mail folders (Inbox, Drafts, Sent Items, etc.).
- **create_email_rule**: Create a mail rule with conditions and actions.

### Category, profile, and settings tools

- **create_master_category**: Create a new category in the user’s master category list.
- **get_master_categories**: List all master categories.
- **get_profile**: Get the user’s Outlook profile (basic user information).
- **get_mailbox_settings**: Read mailbox settings (automatic replies, time zone, working hours, etc.).
- **update_mailbox_settings**: Update automatic replies, language, time zone, or working hours.
- **get_mail_delta**: Get delta changes for messages (for sync scenarios).
- **get_mail_tips**: Get mail tips (automatic replies, mailbox full, etc.) for recipients.
- **get_supported_languages**: List mailbox-supported languages.
- **get_supported_time_zones**: List supported time zones (Windows or IANA).

## Troubleshooting

### "AADSTS65001: The user or administrator has not consented"
- Make sure you've added the required API permissions
- Grant admin consent if possible, or sign in to consent

### "AADSTS700016: Application not found"
- Double-check your `OUTLOOK_CLIENT_ID` in `.env`
- Make sure the app registration exists in Azure

### Token expired
- Normally the token is refreshed silently using the cached refresh token.
- If you see an explicit error like "Authentication expired. Please re-authenticate.":
  - Delete `.token_cache.json` (optional but sometimes helpful).
  - Run `python run_server.py` again and complete the device code login flow.

## License

MIT



