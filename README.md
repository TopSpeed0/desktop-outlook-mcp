# [PowerShell-Outlook-tool](https://github.com/TopSpeed0/desktop-outlook-mcp)

[![GitHub license](https://img.shields.io/github/license/TopSpeed0/PowerShell-Outlook-tool)](https://github.com/TopSpeed0/PowerShell-Outlook-tool/blob/main/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/TopSpeed0/PowerShell-Outlook-tool)](https://github.com/TopSpeed0/PowerShell-Outlook-tool/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/TopSpeed0/PowerShell-Outlook-tool)](https://github.com/TopSpeed0/PowerShell-Outlook-tool/issues)
[![GitHub last commit](https://img.shields.io/github/last-commit/TopSpeed0/PowerShell-Outlook-tool)](https://github.com/TopSpeed0/PowerShell-Outlook-tool/commits/main)
[![GitHub repo size](https://img.shields.io/github/repo-size/TopSpeed0/PowerShell-Outlook-tool)](https://github.com/TopSpeed0/PowerShell-Outlook-tool)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://docs.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)](https://github.com/TopSpeed0/PowerShell-Outlook-tool)
[![Outlook](https://img.shields.io/badge/Outlook-COM%20Automation-0078D4?logo=microsoftoutlook&logoColor=white)](https://github.com/TopSpeed0/PowerShell-Outlook-tool)
[![AI Ready](https://img.shields.io/badge/AI%20Ready-Claude%20%7C%20Copilot-blueviolet?logo=github&logoColor=white)](https://github.com/TopSpeed0/PowerShell-Outlook-tool/blob/main/SKILL.md)

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
| `OutlookTools.psm1` | PowerShell module — `Connect-Outlook`, `Get-OutlookMail`, `Read-OutlookMail`, `Save-OutlookAttachment`, `Send-OutlookReply`, `Send-OutlookMail`, and more |
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

# Download attachments
Save-OutlookAttachment -EntryID $mail.EntryID
Save-OutlookAttachment -EntryID $mail.EntryID -FileNameFilter '\.pdf$' -DestinationPath 'C:\Temp'

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
