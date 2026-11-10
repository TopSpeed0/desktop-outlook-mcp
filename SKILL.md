---
name: outlook-mail
description: Outlook Mail & Calendar AI Skill — read, search, and reply to emails, read and create calendar events, triage unread mail into an Excel review workbook and turn its Reply column into reply drafts, via PowerShell COM. TRIGGER when user mentions Outlook / email / inbox / calendar / meeting / appointment / "check email from X" / "reply to" / "send email" / "what's on my calendar" / -reviewCSV / -SendCSV / -SendCSVRAW / mail review sheet / "what mail needs my action" / Get-OutlookMail / Read-OutlookMail / Send-OutlookReply / Get-OutlookCalendar / New-OutlookAppointment. Requires Outlook desktop running on the machine.
argument-hint: "[-reviewCSV [-Days N] | -SendCSV [file.xlsx] | -SendCSVRAW [file.xlsx]]"
---

# Outlook Mail AI Skill

Read, search, and reply to Outlook emails via PowerShell COM automation. Works with any local Outlook desktop client (Windows).

**Module:** `OutlookTools.psm1` sits next to this `SKILL.md`. If this file is loaded through a
symbolic link, resolve the link target and import the module from that folder.

> **IRON RULE — encrypted mail is readable.** You use the COM interface and act as the user's
> Outlook desktop app, with the same access the user has. Never report that a mail "was
> encrypted and cannot be read".

> **IRON RULE — sending to externals.** Reading encrypted mail is fine, but external
> recipients can only read what you send if the MIP sensitivity label is cleared.
> `Send-OutlookMail` / `Send-OutlookReply` do this automatically; do not use `-KeepLabel`
> unless the user asks.

> **IRON RULE — no sign-off.** Never end a mail or reply with `Thanks,` / `Regards,` /
> `Best,` or the user's name. Outlook adds the user's real signature, so a written one
> shows up twice. End on the last real sentence. `Send-OutlookMail` and `Send-OutlookReply` strip a trailing
> sign-off when Outlook has already added the signature (`-KeepSignOff` turns that off),
> but write the body without one in the first place.

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
| `Set-OutlookMeetingResponse -EntryID -Response <Accept\|Tentative\|Decline> [-Comment] [-Start] [-Send] [-NoResponse]` | Respond to a meeting invitation (alias `Respond-OutlookMeeting`). Accepts pipeline from `Get-OutlookCalendar`. Opens the response for review by default; `-Send` sends it to the organizer; `-NoResponse` updates the calendar only. Supports `-WhatIf`. |
| `Read-OutlookMail -EntryID <id> [-IncludeHTML] [-AsMarkdown] [-MaxBodyLength <int>]` | Full email: body (text), attachments, CC. Use `-AsMarkdown` for clean Markdown output (best for AI). Use `-IncludeHTML` only when you need HTML (e.g. for reply). Use `-MaxBodyLength 2000` to truncate long bodies. |
| `Set-OutlookMailRead -EntryID <id> [-Unread]` | Mark emails as read (or unread with `-Unread`). Accepts pipeline from `Get-OutlookMail`. Supports `-WhatIf`. |
| `Save-OutlookAttachment -EntryID <id> [-DestinationPath] [-FileNameFilter]` | Save attachments to disk. Default destination: `~/Downloads`. Filter by regex (e.g. `'\.pdf$'`). |
| `Send-OutlookReply -EntryID <id> -Body <html> [-ReplyAll] [-Attachments] [-Send] [-Unencrypted] [-KeepLabel] [-KeepSignOff]` | Reply to an email. Opens draft by default; `-Send` sends immediately. Auto-clears the MIP label when a recipient is external. The body goes inside the reply's `<body>`, above Outlook's signature; a trailing sign-off is removed when that signature is present (`SignOffRemoved` in the result). |
| `Send-OutlookMail -To <addr> -Subject <text> -Body <text> [-CC] [-BCC] [-Attachments] [-HTML] [-Send] [-Unencrypted] [-KeepLabel] [-KeepSignOff] [-NoSignature]` | Compose new email. Opens draft by default; `-Send` sends immediately. Auto-clears the MIP label when a recipient is external. Adds the user's new-mail signature (the body goes inside `<body>`, above it); `-NoSignature` skips it. Resolves recipients against the address book like the GUI; an unresolved name gives a warning (draft) or blocks `-Send`. Result reports `SignatureAdded`, `SignOffRemoved` and `Unresolved`. |
| `ConvertTo-EmailMarkdown -Html <string>` | Convert Outlook HTML to clean Markdown. Strips MSO/Word CSS bloat, converts tables/bold/italic/links/lists to Markdown syntax, decodes HTML entities. Pipeline: `Read-OutlookMail -EntryID $id -IncludeHTML \| ConvertTo-EmailMarkdown` |
| `Save-OutlookMail -EntryID <id> -Format <MSG\|HTML\|TXT\|Markdown> [-DestinationPath]` | Save email to file. MSG = lossless, HTML = raw, TXT = plain text, Markdown = clean .md with metadata header. Default: `~/Downloads`. |
| `Get-OutlookReviewCandidates [-Days] [-MaxItems] [-MaxBodyLength] [-IncludeHandled]` | Mail addressed to the user (To or CC) for triage. Default: unread Inbox; `-Days N` = read and unread from the last N days. Skips mail the user sent, and mail already closed (Ignore/Done/False Alarm) or drafted in an earlier workbook unless `-IncludeHandled`. Returns EntryID, Date, From, Subject, Role (To/CC), ToCount, CcCount, Importance, Body (newest message only, quoted thread cut). Does not mark mail read. |
| `Export-OutlookReview [-Path] [-Show]` | Pipe triaged rows in (EntryID, Date, From, Subject, Role + Priority, Category, Summary, ActionNeeded). Writes a formatted `.xlsx` to `~/Downloads/OutlookReview/OutlookReview-<time>.xlsx`, sorted High→Low. |
| `Import-OutlookReview [-Path] [-Pending]` | Read a review workbook (newest one when `-Path` is omitted). `-Pending` = rows with a Reply, no `Drafted` status, and Decision empty or `Reply`. |
| `Send-OutlookReviewReplies [-Path] (-Raw \| -HtmlByEntryID <hashtable>)` | Open a reply **draft** per pending row (never sends), Reply-All when the row says `ReplyAll = Yes`, then writes `Drafted <time>` in Status. `-Raw` uses the Reply text exactly as typed. Refuses to run while the file is open in Excel. |
| `ConvertTo-OutlookReplyHtml -Text <string>` | Plain text → HTML changing only line breaks (blank line = paragraph, newline = `<br>`, double spaces kept). |

## Review Workbook — `-reviewCSV`, `-SendCSV`, `-SendCSVRAW`

When the skill is invoked with one of these arguments (`/outlook-mail -reviewCSV`), run that
workflow. The file is an Excel `.xlsx` despite the "CSV" names. **Replies are always drafts —
never pass `-Send`, never send anything in these workflows.** Needs the `ImportExcel` module.

### `-reviewCSV` — build the review workbook

1. `Get-OutlookReviewCandidates` (add `-Days N` if the user gave a number of days). Save the
   result to a JSON file in your scratchpad and read the bodies from there.
2. Triage every mail yourself. **Keep** a mail when the user must do something (a question to
   them, a request, an approval, an incident they were added to, a security finding in their
   own repo) or it tells them something important (a decision, a plan affecting their systems,
   a completed request they must verify). **Drop** marketing, newsletters, webinars,
   auto-generated notices that need nothing, "thanks"/"works now" replies, and mail with only
   a signature. In a thread with several unread replies, keep the one that carries the
   request and mention the rest in its Summary.
3. For each kept mail set:
   - `Priority`: `High` (security, outage, a direct question blocking someone, a deadline),
     `Medium` (needs action soon or affects the user's systems), `Low` (nice to know / small task).
   - `Category`: `Action` or `Info`.
   - `Summary`: 1–2 sentences — who wants what, with the ticket/host/repo names.
   - `ActionNeeded`: what the user should do, imperative; empty for pure Info.
4. Pipe the rows (with the candidate's EntryID, Date, From, Subject, Role) to
   `Export-OutlookReview -Show`. Report the path, how many mails were scanned/kept, and the
   High items in one short list.

### `-SendCSV` — AI-written replies from the Reply column

1. `Import-OutlookReview -Pending` (pass the path if the user gave one; default = newest file).
2. For each row, read the mail (`Read-OutlookMail -EntryID ... -AsMarkdown -MaxBodyLength 4000`)
   and turn the user's Reply notes into a proper reply: same meaning and decisions, nothing
   invented, follow **Email Styling** and the no-sign-off rule. Keep it as short as the notes.
3. Build `@{ <EntryID> = '<html>' }` for all rows and call
   `Send-OutlookReviewReplies -HtmlByEntryID $html`. Report the drafts opened and any row skipped.

### `-SendCSVRAW` — the user's words, unchanged

`Send-OutlookReviewReplies -Raw`. Do not rewrite, translate, correct or style the text — the
module only converts line breaks to HTML and keeps any sign-off the user wrote.

### Workbook columns

`No, Priority, Category, Date, From, Subject, Role, Summary, ActionNeeded, Decision, Reply,
ReplyAll, Notes, Status, EntryID` (EntryID hidden). The user fills the highlighted columns:

| Column | Values | Effect |
|---|---|---|
| **Decision** | `Reply` / `Ignore` / `Done` / `False Alarm` (drop-down) | `Ignore`, `Done`, `False Alarm` = never drafted (even with Reply text), row greyed out, and the mail is **not shown again** by the next `-reviewCSV`. Empty or `Reply` = drafted when Reply has text. |
| **Reply** | free text | What to answer. Required for a draft. |
| **ReplyAll** | `Yes` / `No` | `Yes` = Reply-All. |
| **Notes** | free text | The user's own record ("false alarm – test secret"). Never sent. Read it when the user asks about a review. |
| **Status** | written by the tool | `Drafted <time>` after the draft opens; that row is skipped next time and the mail is not shown again. Clear it to redo the row. |

The next `-reviewCSV` reads all earlier workbooks in `~/Downloads/OutlookReview` and skips
mail already closed or drafted there; say how many were skipped. `-IncludeHandled` shows them
again. If a file is open in Excel the send step stops before creating anything — ask the user to
close it.

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
- **No sign-off — ever.** See the iron rule at the top. This applies to new mails and
  replies alike.

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
- `-WhatIf` (or declining `-Confirm`) discards the item and leaves **no draft behind**, with
  or without `-Send`. It returns `Status = 'NotSent'` and `Encryption = 'Would be: ...'` so
  you can see what the send would have done. The label is never actually modified on that path.
- The AI should always show the user what it intends to send before using `-Send`.

## Components

- `OutlookTools.psm1` — the PowerShell module (all functions)
- `Restore-MSG.ps1` — original GUI tool for bulk .msg restore (standalone, not part of the AI skill)
- `outlook-config.json` — local mailbox config (gitignored)
- `outlook-config.example.json` — template for new users
