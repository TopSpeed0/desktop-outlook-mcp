# Outlook Mail AI Skill

Read, search, and reply to Outlook emails via PowerShell COM automation. Works with any local Outlook desktop client (Windows).

## Setup

1. Clone this repo
2. Copy `outlook-config.example.json` to `outlook-config.json` and set your mailbox address
3. Outlook must be running on the machine

## One-prompt install

```
Install-OutlookSkill.ps1
```

Or paste this into Claude Code:

```
Import-Module <path-to>\OutlookTools.psm1; Connect-Outlook
```

## Available Functions

| Function | Description |
|---|---|
| `Connect-Outlook` | Attach to running Outlook or launch it. Auto-called by other functions. |
| `Disconnect-Outlook` | Release the COM object. |
| `Get-OutlookProfile` | List all mailboxes/profiles configured in Outlook. |
| `Get-OutlookFolder [-Mailbox] [-FolderPath]` | Browse folders. Shows name, item count, unread count. |
| `Get-OutlookMail [-Mailbox] [-FolderPath] [-Count] [-UnreadOnly] [-From] [-Subject]` | List emails with filters. Returns Index, EntryID, Subject, From, Date, preview. Plain-text filters are pushed down to MAPI via `Restrict()`; regex filters fall back to a client-side walk. |
| `Get-OutlookCalendar [-Mailbox] [-Start] [-Days] [-Count] [-IncludeBody]` | Read appointments and meetings. Default: next 7 days, 30 items. Expands recurring series. Returns Subject, Start, End, Duration, Location, Organizer, IsRecurring, AllDayEvent, BusyStatus, Attendees, EntryID. |
| `New-OutlookAppointment -Subject -Start [-Minutes] [-Location] [-Body] [-Attendees] [-AllDay] [-BusyStatus] [-ReminderMinutes] [-Save] [-Send]` | Create an appointment, or a meeting when `-Attendees` is given. Opens for review by default; `-Save` commits silently, `-Send` dispatches invitations. |
| `Read-OutlookMail -EntryID <id> [-IncludeHTML] [-AsMarkdown] [-MaxBodyLength <int>]` | Full email: body (text), attachments, CC. Use `-AsMarkdown` for clean Markdown output (best for AI). Use `-IncludeHTML` only when you need HTML (e.g. for reply). Use `-MaxBodyLength 2000` to truncate long bodies. |
| `Save-OutlookAttachment -EntryID <id> [-DestinationPath] [-FileNameFilter]` | Save attachments to disk. Default destination: `~/Downloads`. Filter by regex (e.g. `'\.pdf$'`). |
| `Send-OutlookReply -EntryID <id> -Body <html> [-ReplyAll] [-Attachments] [-Send] [-Unencrypted] [-KeepLabel]` | Reply to an email. Opens draft by default; `-Send` sends immediately. Auto-clears the MIP label when a recipient is external. |
| `Send-OutlookMail -To <addr> -Subject <text> -Body <text> [-CC] [-BCC] [-Attachments] [-HTML] [-Send] [-Unencrypted] [-KeepLabel]` | Compose new email. Opens draft by default; `-Send` sends immediately. Auto-clears the MIP label when a recipient is external. |
| `ConvertTo-EmailMarkdown -Html <string>` | Convert Outlook HTML to clean Markdown. Strips MSO/Word CSS bloat, converts tables/bold/italic/links/lists to Markdown syntax, decodes HTML entities. Pipeline: `Read-OutlookMail -EntryID $id -IncludeHTML \| ConvertTo-EmailMarkdown` |
| `Save-OutlookMail -EntryID <id> -Format <MSG\|HTML\|TXT\|Markdown> [-DestinationPath]` | Save email to file. MSG = lossless, HTML = raw, TXT = plain text, Markdown = clean .md with metadata header. Default: `~/Downloads`. |

## Performance

COM objects do NOT persist between PowerShell calls. Each call re-imports the module and reconnects to Outlook. To avoid slowness:

- **Chain everything in a single PowerShell call.** Import once, then run all the commands you need in one shot.
- **Never make separate calls** for find → read → save when you can combine them.

### Token Optimization (AI agents)

- `Read-OutlookMail` now returns **text body only** by default (no HTMLBody). Outlook HTML can be 10-100KB of Word/MSO bloat.
- Use `-AsMarkdown` for the cleanest AI-friendly output — strips all MSO bloat and returns structured Markdown.
- Use `-MaxBodyLength 2000` to cap long email bodies (shows char count so you know if truncated).
- Only add `-IncludeHTML` when you actually need the HTML (e.g., before composing a styled reply).
- Use `Save-OutlookMail -Format Markdown` to save emails as `.md` files with metadata headers.
- `Get-OutlookMail` already limits preview to 200 chars — use it for scanning, then `Read-OutlookMail` only for emails you need.

## AI Usage Pattern

When Claude (or any AI agent) needs to work with email:

```powershell
# Load the module
Import-Module .\OutlookTools.psm1

# Browse inbox (mailbox auto-loaded from outlook-config.json)
Get-OutlookMail -Count 5

# Read a specific email
$mail = Get-OutlookMail -Count 1
Read-OutlookMail -EntryID $mail.EntryID

# Read with body truncation (saves tokens for long emails)
Read-OutlookMail -EntryID $mail.EntryID -MaxBodyLength 2000

# Read with HTML (only when needed for reply formatting)
Read-OutlookMail -EntryID $mail.EntryID -IncludeHTML

# Search by sender or subject
Get-OutlookMail -From 'support@vendor.com' -Count 10
Get-OutlookMail -Subject 'urgent' -UnreadOnly

# Reply (draft — user reviews before sending)
Send-OutlookReply -EntryID $mail.EntryID -Body '<p>Thank you for the update.</p>'

# Reply and send immediately
Send-OutlookReply -EntryID $mail.EntryID -Body '<p>Confirmed, thanks.</p>' -Send

# Download all attachments from an email
Save-OutlookAttachment -EntryID $mail.EntryID

# Download only PDFs to a specific folder
Save-OutlookAttachment -EntryID $mail.EntryID -FileNameFilter '\.pdf$' -DestinationPath 'C:\Temp'

# Download only images
Save-OutlookAttachment -EntryID $mail.EntryID -FileNameFilter '\.(png|jpg|jpeg|gif|bmp)$'

# Read as clean Markdown (best for AI consumption)
Read-OutlookMail -EntryID $mail.EntryID -AsMarkdown

# Read as Markdown with body truncation
Read-OutlookMail -EntryID $mail.EntryID -AsMarkdown -MaxBodyLength 3000

# Convert HTML to Markdown via pipeline
Read-OutlookMail -EntryID $mail.EntryID -IncludeHTML | ConvertTo-EmailMarkdown

# Convert raw HTML string to Markdown
ConvertTo-EmailMarkdown -Html $item.HTMLBody

# Save email to file (various formats)
Save-OutlookMail -EntryID $mail.EntryID -Format MSG                        # lossless .msg
Save-OutlookMail -EntryID $mail.EntryID -Format HTML                       # raw HTML
Save-OutlookMail -EntryID $mail.EntryID -Format TXT                        # plain text
Save-OutlookMail -EntryID $mail.EntryID -Format Markdown                   # clean .md file
Save-OutlookMail -EntryID $mail.EntryID -Format Markdown -DestinationPath C:\Temp

# New email
Send-OutlookMail -To 'someone@company.com' -Subject 'Report' -Body '<b>Attached.</b>' -HTML -Attachments 'C:\report.pdf' -Send

# External recipient — label cleared automatically so they can read it
Send-OutlookMail -To 'partner@vendor.com' -Subject 'Specs' -Body 'See attached.' -Send

# Force unencrypted / force keep the label
Send-OutlookMail -To 'partner@vendor.com' -Subject 'Specs' -Body '...' -Unencrypted -Send
Send-OutlookMail -To 'partner@vendor.com' -Subject 'Confidential' -Body '...' -KeepLabel -Send

# Calendar — next 7 days
Get-OutlookCalendar

# Calendar — next 30 days, up to 100 items, including bodies
Get-OutlookCalendar -Days 30 -Count 100 -IncludeBody

# Calendar — look back over the past week
Get-OutlookCalendar -Start (Get-Date).AddDays(-7) -Days 7

# Create an appointment (opens for review)
New-OutlookAppointment -Subject 'Maintenance window' -Start '2026-08-03 22:00' -Minutes 120 -Location 'DC1'

# Create and save silently
New-OutlookAppointment -Subject 'Focus block' -Start '2026-08-04 09:00' -Minutes 90 -BusyStatus Busy -Save

# Send a meeting invitation
New-OutlookAppointment -Subject 'Design review' -Start '2026-08-05 10:00' -Attendees 'alice@company.com','bob@company.com' -Send
```

**IMPORTANT:** Do NOT call Outlook COM methods directly (e.g. `$item.GetInspector`, `$namespace.GetItemFromID`). Always use the module functions — they handle connection, mailbox resolution, and cleanup.

## Config

`outlook-config.json` is loaded automatically at import time from the module directory via `$PSScriptRoot` (the folder where the `.psm1` lives — not your current working directory).

```json
{
  "mailbox": "Your.Name@company.com"
}
```

To create from template:
```powershell
Copy-Item outlook-config.example.json outlook-config.json
# Edit outlook-config.json and set your real mailbox address
```

When `mailbox` is set, all functions use it as default — no need to pass `-Mailbox` every time. If the config is missing or has the wrong mailbox, you'll get: `No mailbox specified`.

## Email Styling — ALWAYS Apply

When composing or replying to emails, always use professional styled HTML:
- Font: `Calibri, Arial, sans-serif` 11pt, color `#333`
- **Green success banners** for completed items / good news: `background: #e8f5e9; border-left: 4px solid #43a047; padding: 10px 14px; border-radius: 4px`
- **Clean table layouts** for structured point-by-point responses: bold label column (width ~140px, color `#555`) + content column, separated by `border-bottom: 1px solid #e0e0e0`
- Nested mini-tables for data breakdowns (version lists, VM counts, etc.)
- `<a href="mailto:...">@Name</a>` for @-mentions
- Inline styles only — Outlook ignores `<style>` blocks
- No Word/MSO bloat — keep HTML clean and minimal
- Tone: professional but concise, not overly formal
- **No sign-off — ever.** Do not end with `Regards,` / `Thanks,` / `Best,` or the
  sender's name. Outlook appends the user's real signature on send, so a written one
  duplicates it. End on the last substantive sentence. This applies to replies too.

## Sensitivity Labels & External Recipients

Outlook applies a MIP sensitivity label on send, and if that label has IRM enabled the
message is encrypted. **External recipients cannot read it.** Both send functions handle
this automatically:

- **Internal only** (domains in `internalDomains`) — label left in place, mail is encrypted.
- **Any external recipient** — the label is cleared and `Permission = 0` set, so the mail
  goes out readable. A yellow warning names the external addresses.
- `-Unencrypted` — force the label to be cleared regardless of recipients.
- `-KeepLabel` — force the label to be kept, even for external recipients.

Configure the internal domain list in `outlook-config.json`:

```json
{
  "mailbox": "Your.Name@company.com",
  "internalDomains": ["company.com", "subsidiary.com"]
}
```

Subdomains of an internal domain count as internal (`mail.company.com` matches
`company.com`), but lookalikes do not (`notcompany.com` is external).

**`internalDomains` is empty by default.** Until you configure it, no recipient is
classified as external and the label is always kept — a warning is emitted once per
session. That direction is deliberate: an unconfigured install must not silently strip
protection from every message you send. Use `-Unencrypted` per message in the meantime.

**When an address cannot be resolved to SMTP, it is treated as internal** — the label is
kept. Keeping a label unnecessarily only makes the mail unreadable; clearing one
unnecessarily could expose protected content. The failure mode is deliberate.

Both functions return `Encryption` (`Cleared` or `LabelKept`) and `ExternalRecipients`
so the caller can confirm what happened.

## Performance Notes

`Get-OutlookMail` translates plain-text `-Subject`, `-From`, and `-UnreadOnly` filters into
a DASL `@SQL=` query and pushes them into MAPI via `Items.Restrict()` before walking the
folder. On a large mailbox this is the difference between filtering thousands of items
in Outlook and marshalling them all across COM.

Filters containing regex metacharacters (`^ $ . | ? * + ( ) [ ] { } \`) are **not** pushed
down — DASL `LIKE` would interpret them literally and change the result. Those fall back
to the client-side walk, so regex support is unchanged. Any `Restrict()` failure also
falls back silently. Use `-Verbose` to see which path ran.

`-From` matches sender display name **or** SMTP address in both paths.

## Safety

- `Send-OutlookReply` and `Send-OutlookMail` open a **draft** by default. Add `-Send` to send immediately.
- `New-OutlookAppointment` opens for review by default. `-Save` commits, `-Send` invites attendees.
- All support `-WhatIf` and `-Confirm` via `SupportsShouldProcess`.
- `-WhatIf` (or declining `-Confirm`) discards the item and leaves **no draft behind**. It
  returns `Status = 'NotSent'` and `Encryption = 'Would be: ...'` so you can see what the
  send would have done. The label is never actually modified on that path.
- The AI should always show the user what it intends to send before using `-Send`.

## Components

- `OutlookTools.psm1` — the PowerShell module (all functions)
- `Restore-MSG.ps1` — original GUI tool for bulk .msg restore (standalone, not part of the AI skill)
- `outlook-config.json` — local mailbox config (gitignored)
- `outlook-config.example.json` — template for new users
