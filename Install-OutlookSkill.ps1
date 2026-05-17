# Install-OutlookSkill.ps1 — One-prompt setup for the Outlook Mail AI Skill.
#
# Usage:
#   irm https://raw.githubusercontent.com/TopSpeed0/PowerShell-Outlook-tool/main/Install-OutlookSkill.ps1 | iex
#   — or —
#   .\Install-OutlookSkill.ps1
#   — or —
#   .\Install-OutlookSkill.ps1 -Mailbox 'First.Last@company.com'

param(
    [string]$Mailbox,
    [string]$InstallPath = "$env:USERPROFILE\.claude\skills\outlook-mail"
)

$ErrorActionPreference = 'Stop'
$repoUrl = 'https://github.com/TopSpeed0/PowerShell-Outlook-tool.git'

Write-Host '=== Outlook Mail AI Skill Installer ===' -ForegroundColor Cyan

# 1. Clone or update
if (Test-Path (Join-Path $InstallPath '.git')) {
    Write-Host "Updating existing install at $InstallPath ..."
    git -C $InstallPath pull --ff-only
}
elseif (Test-Path $InstallPath) {
    Write-Host "Directory exists but is not a git repo. Pulling fresh ..."
    Remove-Item $InstallPath -Recurse -Force
    git clone $repoUrl $InstallPath
}
else {
    Write-Host "Cloning to $InstallPath ..."
    New-Item -Path (Split-Path $InstallPath) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    git clone $repoUrl $InstallPath
}

# 2. Create local config
$cfgPath = Join-Path $InstallPath 'outlook-config.json'
if (-not (Test-Path $cfgPath)) {
    if (-not $Mailbox) {
        $profiles = @()
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            $ns = $outlook.GetNamespace('MAPI')
            $profiles = @($ns.Session.Folders | ForEach-Object { $_.Name })
        }
        catch {
            Write-Host 'Outlook not running — cannot auto-detect mailbox.' -ForegroundColor Yellow
        }

        if ($profiles.Count -eq 1) {
            $Mailbox = $profiles[0]
            Write-Host "Auto-detected mailbox: $Mailbox" -ForegroundColor Green
        }
        elseif ($profiles.Count -gt 1) {
            Write-Host 'Multiple mailboxes found:' -ForegroundColor Yellow
            for ($i = 0; $i -lt $profiles.Count; $i++) { Write-Host "  [$i] $($profiles[$i])" }
            $idx = Read-Host 'Select mailbox number'
            $Mailbox = $profiles[[int]$idx]
        }
        else {
            $Mailbox = Read-Host 'Enter your mailbox address (e.g. First.Last@company.com)'
        }
    }
    @{ mailbox = $Mailbox } | ConvertTo-Json | Set-Content $cfgPath -Encoding UTF8
    Write-Host "Config saved: $cfgPath" -ForegroundColor Green
}
else {
    $existing = (Get-Content $cfgPath -Raw | ConvertFrom-Json).mailbox
    Write-Host "Config exists: $existing" -ForegroundColor Green
}

# 3. Verify
Write-Host "`nTesting module load ..." -ForegroundColor Cyan
Import-Module (Join-Path $InstallPath 'OutlookTools.psm1') -Force
Connect-Outlook | Out-Null
$mb = (Get-Content $cfgPath -Raw | ConvertFrom-Json).mailbox
$count = (Get-OutlookMail -Mailbox $mb -Count 1 | Measure-Object).Count
Write-Host "Module loaded. Found $count mail item(s) in Inbox." -ForegroundColor Green

Write-Host @"

=== Installation Complete ===

To use in Claude Code or PowerShell:

  Import-Module '$InstallPath\OutlookTools.psm1'
  Get-OutlookMail -Count 5

Skill reference: $InstallPath\SKILL.md
"@ -ForegroundColor Cyan
