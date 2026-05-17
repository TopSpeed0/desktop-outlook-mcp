# PowerShell-Outlook-tool

PowerShell toolkit for Outlook desktop automation via COM. AI-ready module for Claude Code / GitHub Copilot.

## Quick Start

### One-liner install (for AI agents)

```powershell
irm https://raw.githubusercontent.com/TopSpeed0/PowerShell-Outlook-tool/main/Install-OutlookSkill.ps1 | iex
```

### Manual setup

```powershell
git clone https://github.com/TopSpeed0/PowerShell-Outlook-tool.git
cd PowerShell-Outlook-tool
Copy-Item outlook-config.example.json outlook-config.json
# Edit outlook-config.json — set your mailbox address
```

## What's Inside

| File | Description |
|---|---|
| `OutlookTools.psm1` | PowerShell module — `Connect-Outlook`, `Get-OutlookMail`, `Read-OutlookMail`, `Send-OutlookReply`, `Send-OutlookMail`, and more |
| `Install-OutlookSkill.ps1` | One-prompt installer — clones, configures, and verifies |
| `SKILL.md` | AI skill reference (function docs, usage patterns, safety notes) |

## Usage

```powershell
Import-Module .\OutlookTools.psm1

# List recent emails
Get-OutlookMail -Count 5

# Search by sender
Get-OutlookMail -From 'someone@company.com' -Count 10

# Read full email
$mail = Get-OutlookMail -Count 1
Read-OutlookMail -EntryID $mail.EntryID

# Reply (opens draft by default)
Send-OutlookReply -EntryID $mail.EntryID -Body '<p>Thanks!</p>'

# Send new email
Send-OutlookMail -To 'user@company.com' -Subject 'Report' -Body '<b>See attached.</b>' -HTML -Attachments 'C:\report.pdf' -Send
```

## Config

Create `outlook-config.json` (gitignored) from the example:

```json
{
  "mailbox": "Your.Name@company.com"
}
```

All functions use this as the default mailbox.

## Requirements

- Windows with Outlook desktop installed and running
- PowerShell 5.1+ (or PowerShell 7+)

## Safety

- `Send-OutlookReply` and `Send-OutlookMail` open a **draft** by default — add `-Send` to send immediately
- Both support `-WhatIf` and `-Confirm`

## License

MIT
