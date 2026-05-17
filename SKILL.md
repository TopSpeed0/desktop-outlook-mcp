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
| `Get-OutlookMail [-Mailbox] [-FolderPath] [-Count] [-UnreadOnly] [-From] [-Subject]` | List emails with filters. Returns Index, EntryID, Subject, From, Date, preview. |
| `Read-OutlookMail -EntryID <id>` | Full email: body, HTML, attachments, CC, all fields. |
| `Send-OutlookReply -EntryID <id> -Body <html> [-ReplyAll] [-Send]` | Reply to an email. Opens draft by default; `-Send` sends immediately. |
| `Send-OutlookMail -To <addr> -Subject <text> -Body <text> [-CC] [-Attachments] [-HTML] [-Send]` | Compose new email. Opens draft by default; `-Send` sends immediately. |

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

# Search by sender or subject
Get-OutlookMail -From 'support@vendor.com' -Count 10
Get-OutlookMail -Subject 'urgent' -UnreadOnly

# Reply (draft — user reviews before sending)
Send-OutlookReply -EntryID $mail.EntryID -Body '<p>Thank you for the update.</p>'

# Reply and send immediately
Send-OutlookReply -EntryID $mail.EntryID -Body '<p>Confirmed, thanks.</p>' -Send

# New email
Send-OutlookMail -To 'someone@company.com' -Subject 'Report' -Body '<b>Attached.</b>' -HTML -Attachments 'C:\report.pdf' -Send
```

## Config

`outlook-config.json` (gitignored):
```json
{
  "mailbox": "Your.Name@company.com"
}
```

When `mailbox` is set, all functions use it as default — no need to pass `-Mailbox` every time.

## Safety

- `Send-OutlookReply` and `Send-OutlookMail` open a **draft** by default. Add `-Send` to send immediately.
- Both support `-WhatIf` and `-Confirm` via `SupportsShouldProcess`.
- The AI should always show the user what it intends to send before using `-Send`.

## Components

- `OutlookTools.psm1` — the PowerShell module (all functions)
- `Restore-MSG.ps1` — original GUI tool for bulk .msg restore (standalone, not part of the AI skill)
- `outlook-config.json` — local mailbox config (gitignored)
- `outlook-config.example.json` — template for new users
