# [desktop-outlook-mcp](https://github.com/TopSpeed0/desktop-outlook-mcp)

[![GitHub license](https://img.shields.io/github/license/TopSpeed0/desktop-outlook-mcp)](https://github.com/TopSpeed0/desktop-outlook-mcp/blob/main/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/TopSpeed0/desktop-outlook-mcp)](https://github.com/TopSpeed0/desktop-outlook-mcp/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/TopSpeed0/desktop-outlook-mcp)](https://github.com/TopSpeed0/desktop-outlook-mcp/issues)
[![GitHub last commit](https://img.shields.io/github/last-commit/TopSpeed0/desktop-outlook-mcp)](https://github.com/TopSpeed0/desktop-outlook-mcp/commits/main)
[![GitHub repo size](https://img.shields.io/github/repo-size/TopSpeed0/desktop-outlook-mcp)](https://github.com/TopSpeed0/desktop-outlook-mcp)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://docs.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)](https://github.com/TopSpeed0/desktop-outlook-mcp)
[![Outlook](https://img.shields.io/badge/Outlook-COM%20Automation-0078D4?logo=microsoftoutlook&logoColor=white)](https://github.com/TopSpeed0/desktop-outlook-mcp)
[![AI Ready](https://img.shields.io/badge/AI%20Ready-Claude%20Code%20%7C%20Claude%20Desktop%20%7C%20VS%20Code-blueviolet?logo=github&logoColor=white)](https://github.com/TopSpeed0/desktop-outlook-mcp/blob/main/SKILL.md)

Zero-Config, local-only MCP-ready Outlook automation for Windows.

Built for:

- Claude Code
- Claude Desktop
- VS Code AI tools (Cline, Roo Code, and similar MCP clients)

This repository provides a local PowerShell tool layer for Outlook COM automation with safe defaults, one-command setup, and no cloud relay.

## Quick Start

### One-liner install (recommended for users and AI agents)

```powershell
irm https://raw.githubusercontent.com/TopSpeed0/desktop-outlook-mcp/main/Install-OutlookSkill.ps1 | iex
```

The installer clones the repo, creates local config, and verifies module access.

### Manual setup

```powershell
git clone https://github.com/TopSpeed0/desktop-outlook-mcp.git
cd desktop-outlook-mcp
Copy-Item outlook-config.example.json outlook-config.json
# Edit outlook-config.json with your mailbox address
```

## Config and Security

`outlook-config.json` is intentionally local-only and safely ignored by git.

- It is listed in `.gitignore` so mailbox values are not committed.
- The one-liner setup creates it on your machine during install.
- You can keep personal mailbox settings private while sharing the repo safely.

Example config:

```json
{
  "mailbox": "Your.Name@company.com"
}
```

## Instructions for AI Agents

Copy and paste this into Claude Code, Claude Desktop guidance, or your VS Code AI agent:

```text
Install and register desktop-outlook-mcp from https://github.com/TopSpeed0/desktop-outlook-mcp.

Steps:
1) Run:
   irm https://raw.githubusercontent.com/TopSpeed0/desktop-outlook-mcp/main/Install-OutlookSkill.ps1 | iex
2) Ensure the tool folder exists at:
   %USERPROFILE%\.claude\skills\outlook-mail
3) Register this local MCP server in mcpServers using PowerShell stdio:
   command: pwsh
   args: -NoLogo -NoProfile -Command "Import-Module '$env:USERPROFILE\.claude\skills\outlook-mail\OutlookTools.psm1'; Connect-Outlook | Out-Null; Write-Output 'desktop-outlook-mcp ready'"
4) Use SKILL.md in that folder for available commands and examples.
```

Example `mcpServers` entry:

```json
{
  "mcpServers": {
    "desktop-outlook-mcp": {
      "command": "pwsh",
      "args": [
        "-NoLogo",
        "-NoProfile",
        "-Command",
        "Import-Module '$env:USERPROFILE\\.claude\\skills\\outlook-mail\\OutlookTools.psm1'; Connect-Outlook | Out-Null"
      ]
    }
  }
}
```

## What's Inside

| File | Description |
|---|---|
| `OutlookTools.psm1` | PowerShell module with `Connect-Outlook`, `Get-OutlookMail`, `Read-OutlookMail`, `Save-OutlookAttachment`, `Send-OutlookReply`, `Send-OutlookMail`, and more |
| `Install-OutlookSkill.ps1` | One-prompt installer that clones, configures, and verifies everything |
| `SKILL.md` | AI-oriented reference with usage patterns, examples, and safety notes |

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

## Requirements

- Windows with Outlook desktop installed and running
- PowerShell 5.1+ (or PowerShell 7+)

## Safety

- `Send-OutlookReply` and `Send-OutlookMail` open a draft by default; use `-Send` only when ready.
- Both functions support `-WhatIf` and `-Confirm`.

## License

MIT
