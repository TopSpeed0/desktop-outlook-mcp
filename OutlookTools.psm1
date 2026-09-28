$script:Outlook = $null
$script:Namespace = $null
$script:ModuleRoot = $PSScriptRoot
$script:DefaultMailbox = $null
# Domains treated as internal for sensitivity-label purposes. Set 'internalDomains'
# in outlook-config.json. When it is not configured, no recipient is classified as
# external and the label is always kept -- see Test-ExternalRecipient.
$script:InternalDomains = @()

# MAPI property holding the MIP (Microsoft Information Protection) sensitivity label.
# A label with IRM enabled encrypts the message; external recipients cannot read it.
$script:MipLabelProperty = 'http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/msip_labels'

# DASL property URIs for MAPI-side filtering via Items.Restrict()
$script:Dasl = @{
    Subject     = 'urn:schemas:httpmail:subject'
    SenderName  = 'urn:schemas:httpmail:sendername'
    SenderEmail = 'urn:schemas:httpmail:fromemail'
    Read        = 'urn:schemas:httpmail:read'
}

$cfgPath = Join-Path $PSScriptRoot 'outlook-config.json'
if (Test-Path $cfgPath) {
    $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
    $script:DefaultMailbox = $cfg.mailbox
    if ($cfg.internalDomains) { $script:InternalDomains = @($cfg.internalDomains) }
}

# ---------------------------------------------------------------------------
# Private helpers (not exported)
# ---------------------------------------------------------------------------

# True when the pattern contains no regex metacharacters, i.e. it is safe to
# translate into a DASL LIKE '%...%' prefilter without changing match semantics.
function Test-PlainTextPattern {
    param([string]$Pattern)
    return $Pattern -notmatch '[\\^$.|?*+()\[\]{}]'
}

function New-DaslLikeClause {
    param([string]$Property, [string]$Value)
    # DASL string literals are single-quoted; embedded quotes are doubled.
    '"' + $Property + '" LIKE ' + "'%" + ($Value -replace "'", "''") + "%'"
}

function New-DaslEqualsClause {
    param([string]$Property, [int]$Value)
    '"' + $Property + '" = ' + $Value
}

# Outlook's Restrict() date literals are locale-sensitive. Always format with
# the invariant culture in the form Outlook parses reliably.
function Format-OutlookDate {
    param([datetime]$Date)
    $Date.ToString('MM/dd/yyyy hh:mm tt', [System.Globalization.CultureInfo]::InvariantCulture)
}

# Resolve a Recipient to an SMTP address. Exchange recipients often expose an
# X500/EX legacy DN instead, in which case we ask the directory for the primary
# SMTP address. Returns $null when it cannot be resolved.
function Resolve-RecipientSmtp {
    param($Recipient)
    $addr = $null
    try { $addr = $Recipient.Address } catch {}
    if ($addr -and $addr -match '@') { return $addr }
    try {
        $entry = $Recipient.AddressEntry
        if ($entry) {
            $user = $entry.GetExchangeUser()
            if ($user -and $user.PrimarySmtpAddress) { return $user.PrimarySmtpAddress }
        }
    }
    catch {}
    return $null
}

# Extract e-mail addresses from a free-text recipient string ("a@b.com; c@d.com").
function Get-AddressesFromString {
    param([string]$Text)
    if (-not $Text) { return @() }
    [regex]::Matches($Text, '[\w.+\-'']+@[\w\-]+(?:\.[\w\-]+)+') | ForEach-Object { $_.Value }
}

# Classify a set of addresses against $script:InternalDomains.
# Unresolvable addresses are treated as INTERNAL on purpose: keeping a label we
# should have cleared only makes the mail unreadable, while clearing a label we
# should have kept could expose protected content. Fail toward protection.
function Test-ExternalRecipient {
    param([string[]]$Addresses)

    # With no internal domain list we cannot tell internal from external. Report no
    # external recipients so the label is kept, rather than stripping protection from
    # every message. Warn once so the misconfiguration is visible.
    if (-not $script:InternalDomains -or $script:InternalDomains.Count -eq 0) {
        if (-not $script:WarnedNoInternalDomains) {
            Write-Warning "No 'internalDomains' configured in outlook-config.json - sensitivity labels will always be kept. External recipients may be unable to read your mail; use -Unencrypted to override per message."
            $script:WarnedNoInternalDomains = $true
        }
        return [PSCustomObject]@{ HasExternal = $false; External = @() }
    }

    $external = @()
    foreach ($a in $Addresses) {
        if (-not $a -or $a -notmatch '@') { continue }
        $domain = ($a -split '@')[-1].Trim().TrimEnd('>').ToLowerInvariant()
        $isInternal = $false
        foreach ($d in $script:InternalDomains) {
            $d = $d.Trim().ToLowerInvariant()
            if ($domain -eq $d -or $domain.EndsWith(".$d")) { $isInternal = $true; break }
        }
        if (-not $isInternal) { $external += $a }
    }
    return [PSCustomObject]@{
        HasExternal = ($external.Count -gt 0)
        External    = $external
    }
}

# Strip the MIP sensitivity label and IRM restriction so external recipients can
# read the message. Save() must run after clearing, otherwise Outlook re-applies
# the policy label during Send().
function Clear-OutlookSensitivityLabel {
    param($MailItem)
    try { $MailItem.Permission = 0 } catch { Write-Warning "Could not set Permission = 0: $($_.Exception.Message)" }
    try {
        $MailItem.PropertyAccessor.SetProperty($script:MipLabelProperty, '')
    }
    catch {
        Write-Warning "Could not clear MIP label: $($_.Exception.Message)"
    }
    try { $MailItem.Save() } catch { Write-Warning "Could not save after clearing label: $($_.Exception.Message)" }
}

# Decide whether the label should be cleared. PURE — mutates nothing, so it is safe
# to call before ShouldProcess and under -WhatIf.
function Get-OutlookEncryptionDecision {
    param(
        [string[]]$Addresses,
        [switch]$Unencrypted,
        [switch]$KeepLabel
    )
    if ($KeepLabel) {
        return [PSCustomObject]@{ ShouldClear = $false; Encryption = 'LabelKept'; ExternalRecipients = @(); Reason = '-KeepLabel specified' }
    }

    $check = Test-ExternalRecipient -Addresses $Addresses

    if ($Unencrypted) {
        return [PSCustomObject]@{ ShouldClear = $true; Encryption = 'Cleared'; ExternalRecipients = $check.External; Reason = '-Unencrypted specified' }
    }
    if ($check.HasExternal) {
        return [PSCustomObject]@{ ShouldClear = $true; Encryption = 'Cleared'; ExternalRecipients = $check.External; Reason = "external recipient(s): $($check.External -join ', ')" }
    }
    return [PSCustomObject]@{ ShouldClear = $false; Encryption = 'LabelKept'; ExternalRecipients = @(); Reason = 'all recipients internal' }
}

# Apply a decision. Call this ONLY once committed to sending or displaying — it calls
# Save() on the item, which would otherwise leave a stray draft behind under -WhatIf.
function Invoke-OutlookEncryptionDecision {
    param($MailItem, $Decision)
    if ($Decision.ShouldClear) {
        Clear-OutlookSensitivityLabel -MailItem $MailItem
        Write-Host "Sensitivity label cleared - $($Decision.Reason). Use -KeepLabel to override." -ForegroundColor Yellow
    }
    else {
        Write-Verbose "Sensitivity label kept - $($Decision.Reason)."
    }
}

# Discard an unsent, unsaved item so -WhatIf leaves no trace. olDiscard = 1.
function Remove-UncommittedOutlookItem {
    param($MailItem)
    try { $MailItem.Close(1) } catch { Write-Verbose "Could not discard item: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------------------

function Connect-Outlook {
    [CmdletBinding()]
    param()
    if ($script:Outlook) { return $script:Outlook }
    try {
        $script:Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    }
    catch {
        $script:Outlook = New-Object -ComObject Outlook.Application
    }
    $script:Namespace = $script:Outlook.GetNamespace('MAPI')
    Write-Host "Connected to Outlook." -ForegroundColor Green
    $script:Outlook
}

function Disconnect-Outlook {
    [CmdletBinding()]
    param()
    if ($script:Namespace) {
        try { $script:Namespace.Logoff() } catch {}
    }
    if ($script:Outlook) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Outlook) | Out-Null } catch {}
    }
    $script:Outlook = $null
    $script:Namespace = $null
    Write-Host "Disconnected from Outlook." -ForegroundColor Yellow
}

function Get-OutlookProfile {
    [CmdletBinding()]
    param()
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $script:Namespace.Session.Folders | ForEach-Object {
        [PSCustomObject]@{
            Name      = $_.Name
            FolderPath = $_.FolderPath
        }
    }
}

function Get-OutlookFolder {
    [CmdletBinding()]
    param(
        [string]$Mailbox = $script:DefaultMailbox,

        [string]$FolderPath
    )
    if (-not $Mailbox) { throw "No mailbox specified. Pass -Mailbox or set mailbox in outlook-config.json." }
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $root = $script:Namespace.Session.Folders.Item($Mailbox)
    if (-not $root) { throw "Mailbox '$Mailbox' not found." }

    if (-not $FolderPath) {
        $root.Folders | ForEach-Object {
            [PSCustomObject]@{
                Name       = $_.Name
                ItemCount  = $_.Items.Count
                UnreadCount = $_.UnReadItemCount
                FolderPath = $_.FolderPath
            }
        }
        return
    }

    $current = $root
    foreach ($part in ($FolderPath -split '\\' | Where-Object { $_ })) {
        $found = $null
        foreach ($f in $current.Folders) {
            if ($f.Name -eq $part) { $found = $f; break }
        }
        if (-not $found) { throw "Folder '$part' not found under '$($current.Name)'." }
        $current = $found
    }

    $current.Folders | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            ItemCount   = $_.Items.Count
            UnreadCount = $_.UnReadItemCount
            FolderPath  = $_.FolderPath
        }
    }
}

function Get-OutlookMail {
    [CmdletBinding()]
    param(
        [string]$Mailbox = $script:DefaultMailbox,

        [string]$FolderPath = 'Inbox',

        [int]$Count = 10,

        [switch]$UnreadOnly,

        [string]$From,

        [string]$Subject
    )
    if (-not $Mailbox) { throw "No mailbox specified. Pass -Mailbox or set mailbox in outlook-config.json." }
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }

    $folder = $script:Namespace.Session.Folders.Item($Mailbox)
    if (-not $folder) { throw "Mailbox '$Mailbox' not found." }

    foreach ($part in ($FolderPath -split '\\' | Where-Object { $_ })) {
        $found = $null
        foreach ($f in $folder.Folders) {
            if ($f.Name -eq $part) { $found = $f; break }
        }
        if (-not $found) { throw "Folder '$part' not found." }
        $folder = $found
    }

    # MAPI-side prefilter. Cuts the collection down inside Outlook before the
    # client-side loop below walks it — the difference is large on big folders.
    #
    # Only plain-text values are translated: -From/-Subject accept regex, and a
    # pattern like '^john' or 'a|b' would mean something different to DASL LIKE.
    # When a regex is supplied we skip that clause and let the loop handle it, so
    # match semantics never change. Any Restrict failure falls back to a full walk.
    $items = $folder.Items
    $clauses = @()
    if ($UnreadOnly) {
        $clauses += New-DaslEqualsClause -Property $script:Dasl.Read -Value 0
    }
    if ($Subject -and (Test-PlainTextPattern $Subject)) {
        $clauses += New-DaslLikeClause -Property $script:Dasl.Subject -Value $Subject
    }
    if ($From -and (Test-PlainTextPattern $From)) {
        # The loop matches SenderName OR SenderEmailAddress, so the prefilter must too.
        $clauses += '(' +
            (New-DaslLikeClause -Property $script:Dasl.SenderName  -Value $From) + ' OR ' +
            (New-DaslLikeClause -Property $script:Dasl.SenderEmail -Value $From) + ')'
    }

    if ($clauses.Count -gt 0) {
        $filter = '@SQL=' + ($clauses -join ' AND ')
        try {
            $items = $folder.Items.Restrict($filter)
            Write-Verbose "DASL prefilter applied: $filter"
        }
        catch {
            Write-Verbose "Restrict() failed, falling back to full folder walk: $($_.Exception.Message)"
            $items = $folder.Items
        }
    }

    try { $items.Sort('[ReceivedTime]', $true) }
    catch { Write-Verbose "Sort failed on this collection: $($_.Exception.Message)" }

    $collected = 0
    foreach ($item in $items) {
        if ($collected -ge $Count) { break }
        if ($item.Class -ne 43) { continue } # 43 = olMail
        if ($UnreadOnly -and $item.UnRead -eq $false) { continue }
        # Encrypted/rights-managed items return null from folder collection — resolve fully
        if (-not $item.Subject -and $item.EntryID) {
            try { $item = $script:Namespace.GetItemFromID($item.EntryID) } catch {}
        }

        if ($From -and $item.SenderEmailAddress -notmatch $From -and $item.SenderName -notmatch $From) { continue }
        if ($Subject -and $item.Subject -notmatch $Subject) { continue }

        [PSCustomObject]@{
            Index        = $collected
            EntryID      = $item.EntryID
            Subject      = $item.Subject
            From         = $item.SenderName
            FromEmail    = $item.SenderEmailAddress
            To           = $item.To
            ReceivedTime = $item.ReceivedTime
            UnRead       = $item.UnRead
            BodyPreview  = if ($item.Body) { ($item.Body -replace '\r?\n', ' ').Substring(0, [Math]::Min(200, ($item.Body -replace '\r?\n', ' ').Length)) } else { '' }
        }
        $collected++
    }
}

function Read-OutlookMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$EntryID,

        [switch]$IncludeHTML,

        [switch]$AsMarkdown,

        [int]$MaxBodyLength = 0
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $item = $script:Namespace.GetItemFromID($EntryID)
    if (-not $item) { throw "Mail item not found." }

    if ($AsMarkdown) {
        $bodyText = ConvertTo-EmailMarkdown -Html $item.HTMLBody
    } else {
        $bodyText = $item.Body
    }

    $fullLength = $bodyText.Length
    if ($MaxBodyLength -gt 0 -and $bodyText.Length -gt $MaxBodyLength) {
        $bodyText = $bodyText.Substring(0, $MaxBodyLength) + "`n`n[...truncated at $MaxBodyLength chars, total $fullLength chars]"
    }

    $result = [ordered]@{
        EntryID      = $item.EntryID
        Subject      = $item.Subject
        From         = $item.SenderName
        FromEmail    = $item.SenderEmailAddress
        To           = $item.To
        CC           = $item.CC
        ReceivedTime = $item.ReceivedTime
        UnRead       = $item.UnRead
        Body         = $bodyText
        BodyLength   = $fullLength
        Attachments  = @($item.Attachments | ForEach-Object { $_.FileName })
    }

    if ($IncludeHTML) {
        $result['HTMLBody'] = $item.HTMLBody
    }

    [PSCustomObject]$result
}

function Save-OutlookAttachment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$EntryID,

        [string]$DestinationPath = (Join-Path $env:USERPROFILE 'Downloads'),

        [string]$FileNameFilter
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $item = $script:Namespace.GetItemFromID($EntryID)
    if (-not $item) { throw "Mail item not found." }
    if ($item.Attachments.Count -eq 0) { Write-Warning "No attachments on this email."; return }

    if (-not (Test-Path $DestinationPath)) {
        New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
    }

    $saved = @()
    foreach ($att in $item.Attachments) {
        if ($FileNameFilter -and $att.FileName -notmatch $FileNameFilter) { continue }
        $dest = Join-Path $DestinationPath $att.FileName
        $att.SaveAsFile($dest)
        $saved += [PSCustomObject]@{
            FileName = $att.FileName
            Size     = $att.Size
            Path     = $dest
        }
        Write-Host "Saved: $dest ($([math]::Round($att.Size / 1KB, 1)) KB)" -ForegroundColor Green
    }

    if ($saved.Count -eq 0) {
        Write-Warning "No attachments matched filter '$FileNameFilter'. Available: $($item.Attachments | ForEach-Object { $_.FileName } | Join-String -Separator ', ')"
    }
    $saved
}

function Send-OutlookReply {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$EntryID,

        [Parameter(Mandatory)]
        [string]$Body,

        [switch]$ReplyAll,

        [switch]$Send,

        [switch]$Unencrypted,

        [switch]$KeepLabel
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $item = $script:Namespace.GetItemFromID($EntryID)
    if (-not $item) { throw "Mail item not found." }

    $reply = if ($ReplyAll) { $item.ReplyAll() } else { $item.Reply() }
    $reply.HTMLBody = $Body + $reply.HTMLBody

    # Collect the actual recipients Outlook put on the reply, resolving Exchange
    # legacy DNs to SMTP so external detection is accurate.
    $addresses = @()
    try {
        foreach ($r in $reply.Recipients) {
            try { $r.Resolve() | Out-Null } catch {}
            $smtp = Resolve-RecipientSmtp -Recipient $r
            if ($smtp) { $addresses += $smtp }
        }
    }
    catch { Write-Verbose "Could not enumerate reply recipients: $($_.Exception.Message)" }
    if (-not $addresses) { $addresses = Get-AddressesFromString -Text $reply.To }

    $decision = Get-OutlookEncryptionDecision -Addresses $addresses `
        -Unencrypted:$Unencrypted -KeepLabel:$KeepLabel

    if ($Send) {
        if ($PSCmdlet.ShouldProcess("Reply to '$($item.Subject)' from $($item.SenderName)", "Send")) {
            Invoke-OutlookEncryptionDecision -MailItem $reply -Decision $decision
            $reply.Send()
            Write-Host "Reply sent to '$($item.Subject)'." -ForegroundColor Green
            return [PSCustomObject]@{ Status = 'Sent'; Subject = $item.Subject; To = $item.SenderName; Recipients = $addresses; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients }
        }
        # -WhatIf / declined at the -Confirm prompt: discard, change nothing.
        Remove-UncommittedOutlookItem -MailItem $reply
        return [PSCustomObject]@{ Status = 'NotSent'; Subject = $item.Subject; To = $item.SenderName; Recipients = $addresses; Encryption = "Would be: $($decision.Encryption)"; ExternalRecipients = $decision.ExternalRecipients }
    }

    Invoke-OutlookEncryptionDecision -MailItem $reply -Decision $decision
    $reply.Display()
    Write-Host "Reply draft opened for '$($item.Subject)'." -ForegroundColor Cyan
    [PSCustomObject]@{ Status = 'Draft'; Subject = $item.Subject; To = $item.SenderName; Recipients = $addresses; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients }
}

function Send-OutlookMail {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$To,

        [Parameter(Mandatory)]
        [string]$Subject,

        [Parameter(Mandatory)]
        [string]$Body,

        [string]$CC,

        [string]$BCC,

        [string[]]$Attachments,

        [switch]$HTML,

        [switch]$Send,

        [switch]$Unencrypted,

        [switch]$KeepLabel
    )
    if (-not $script:Outlook) { Connect-Outlook | Out-Null }

    $mail = $script:Outlook.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    if ($CC) { $mail.CC = $CC }
    if ($BCC) { $mail.BCC = $BCC }
    if ($HTML) { $mail.HTMLBody = $Body } else { $mail.Body = $Body }

    foreach ($att in $Attachments) {
        if (Test-Path $att) { $mail.Attachments.Add($att) | Out-Null }
        else { Write-Warning "Attachment not found: $att" }
    }

    $addresses = @(Get-AddressesFromString -Text $To) +
                 @(Get-AddressesFromString -Text $CC) +
                 @(Get-AddressesFromString -Text $BCC)

    $decision = Get-OutlookEncryptionDecision -Addresses $addresses `
        -Unencrypted:$Unencrypted -KeepLabel:$KeepLabel

    if ($Send) {
        if ($PSCmdlet.ShouldProcess("Send mail '$Subject' to $To", "Send")) {
            Invoke-OutlookEncryptionDecision -MailItem $mail -Decision $decision
            $mail.Send()
            Write-Host "Mail sent: '$Subject' to $To" -ForegroundColor Green
            return [PSCustomObject]@{ Status = 'Sent'; Subject = $Subject; To = $To; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients }
        }
        # -WhatIf / declined at the -Confirm prompt: discard, change nothing.
        Remove-UncommittedOutlookItem -MailItem $mail
        return [PSCustomObject]@{ Status = 'NotSent'; Subject = $Subject; To = $To; Encryption = "Would be: $($decision.Encryption)"; ExternalRecipients = $decision.ExternalRecipients }
    }

    Invoke-OutlookEncryptionDecision -MailItem $mail -Decision $decision
    $mail.Display()
    Write-Host "Draft opened: '$Subject' to $To" -ForegroundColor Cyan
    [PSCustomObject]@{ Status = 'Draft'; Subject = $Subject; To = $To; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients }
}

function Get-OutlookCalendar {
    <#
    .SYNOPSIS
    Read calendar appointments and meetings within a date range.
    .DESCRIPTION
    Resolves the mailbox's default Calendar folder by ID rather than by name, so it
    works on non-English Outlook installs. Expands recurring series.
    .EXAMPLE
    Get-OutlookCalendar
    .EXAMPLE
    Get-OutlookCalendar -Days 30 -Count 100
    .EXAMPLE
    Get-OutlookCalendar -Start (Get-Date).AddDays(-7) -Days 7
    #>
    [CmdletBinding()]
    param(
        [string]$Mailbox = $script:DefaultMailbox,

        [datetime]$Start = (Get-Date),

        [int]$Days = 7,

        [int]$Count = 30,

        [switch]$IncludeBody
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }

    # olFolderCalendar = 9. Resolve via the mailbox's Store so this works for
    # secondary mailboxes and on localised Outlook installs (folder name varies).
    $cal = $null
    if ($Mailbox) {
        try {
            $root = $script:Namespace.Session.Folders.Item($Mailbox)
            if ($root -and $root.Store) { $cal = $root.Store.GetDefaultFolder(9) }
        }
        catch { Write-Verbose "Store.GetDefaultFolder(9) failed for '$Mailbox': $($_.Exception.Message)" }
    }
    if (-not $cal) { $cal = $script:Namespace.GetDefaultFolder(9) }
    if (-not $cal) { throw "Could not resolve the Calendar folder." }

    $items = $cal.Items

    # IncludeRecurrences MUST be set before Sort(), and the collection must be
    # sorted by [Start] ascending, or recurring appointments are silently omitted.
    $items.IncludeRecurrences = $true
    $items.Sort('[Start]')

    $from = Format-OutlookDate -Date $Start
    $to   = Format-OutlookDate -Date $Start.AddDays($Days)
    $filter = "[Start] >= '$from' AND [Start] <= '$to'"

    try {
        $filtered = $items.Restrict($filter)
    }
    catch {
        throw "Calendar Restrict() failed with filter [$filter]: $($_.Exception.Message)"
    }

    $collected = 0
    foreach ($appt in $filtered) {
        if ($collected -ge $Count) { break }
        if ($appt.Class -ne 26) { continue } # 26 = olAppointment

        $result = [ordered]@{
            Subject      = $appt.Subject
            Start        = $appt.Start
            End          = $appt.End
            Duration     = $appt.Duration      # minutes
            Location     = $appt.Location
            Organizer    = $appt.Organizer
            IsRecurring  = $appt.IsRecurring
            AllDayEvent  = $appt.AllDayEvent
            BusyStatus   = switch ($appt.BusyStatus) { 0 { 'Free' } 1 { 'Tentative' } 2 { 'Busy' } 3 { 'OutOfOffice' } 4 { 'WorkingElsewhere' } default { $appt.BusyStatus } }
            Attendees    = $appt.RequiredAttendees
            EntryID      = $appt.EntryID
        }
        if ($IncludeBody) { $result['Body'] = $appt.Body }

        [PSCustomObject]$result
        $collected++
    }
}

function New-OutlookAppointment {
    <#
    .SYNOPSIS
    Create a calendar appointment, or a meeting when -Attendees is supplied.
    .DESCRIPTION
    Opens the item for review by default. Add -Save to commit it silently, or
    -Send to dispatch meeting invitations to attendees.
    .EXAMPLE
    New-OutlookAppointment -Subject 'Maintenance window' -Start '2026-08-03 22:00' -Minutes 120
    .EXAMPLE
    New-OutlookAppointment -Subject 'Design review' -Start '2026-08-03 10:00' -Attendees 'alice@company.com' -Send
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$Subject,

        [Parameter(Mandatory)]
        [datetime]$Start,

        [int]$Minutes = 30,

        [string]$Location,

        [string]$Body,

        [string[]]$Attendees,

        [switch]$AllDay,

        [ValidateSet('Free', 'Tentative', 'Busy', 'OutOfOffice', 'WorkingElsewhere')]
        [string]$BusyStatus = 'Busy',

        [int]$ReminderMinutes = 15,

        [switch]$Save,

        [switch]$Send
    )
    if (-not $script:Outlook) { Connect-Outlook | Out-Null }

    $appt = $script:Outlook.CreateItem(1) # olAppointmentItem
    $appt.Subject = $Subject
    $appt.Start = $Start
    if ($AllDay) { $appt.AllDayEvent = $true } else { $appt.Duration = $Minutes }
    if ($Location) { $appt.Location = $Location }
    if ($Body) { $appt.Body = $Body }
    $appt.BusyStatus = @{ Free = 0; Tentative = 1; Busy = 2; OutOfOffice = 3; WorkingElsewhere = 4 }[$BusyStatus]
    if ($ReminderMinutes -gt 0) {
        $appt.ReminderSet = $true
        $appt.ReminderMinutesBeforeStart = $ReminderMinutes
    }

    if ($Attendees) {
        $appt.MeetingStatus = 1 # olMeeting — required before adding recipients
        foreach ($a in $Attendees) { $appt.Recipients.Add($a) | Out-Null }
        try { $appt.Recipients.ResolveAll() | Out-Null } catch {}
    }

    $when = $Start.ToString('yyyy-MM-dd HH:mm')

    if ($Send -and $Attendees) {
        if ($PSCmdlet.ShouldProcess("Meeting '$Subject' at $when to $($Attendees -join ', ')", "Send invitation")) {
            $appt.Send()
            Write-Host "Meeting invitation sent: '$Subject' at $when" -ForegroundColor Green
            return [PSCustomObject]@{ Status = 'Sent'; Subject = $Subject; Start = $Start; Attendees = $Attendees }
        }
        return
    }

    if ($Save) {
        if ($PSCmdlet.ShouldProcess("Appointment '$Subject' at $when", "Save to calendar")) {
            $appt.Save()
            Write-Host "Appointment saved: '$Subject' at $when" -ForegroundColor Green
            return [PSCustomObject]@{ Status = 'Saved'; Subject = $Subject; Start = $Start; EntryID = $appt.EntryID }
        }
        return
    }

    $appt.Display()
    Write-Host "Appointment draft opened: '$Subject' at $when" -ForegroundColor Cyan
    [PSCustomObject]@{ Status = 'Draft'; Subject = $Subject; Start = $Start }
}

function ConvertTo-EmailMarkdown {
    <#
    .SYNOPSIS
    Converts Outlook HTML email body to clean Markdown for AI consumption.
    .DESCRIPTION
    Strips MSO/Word bloat, converts tables/links/formatting to Markdown.
    Accepts pipeline input from Read-OutlookMail -IncludeHTML.
    .EXAMPLE
    Read-OutlookMail -EntryID $id -IncludeHTML | ConvertTo-EmailMarkdown
    .EXAMPLE
    ConvertTo-EmailMarkdown -Html $item.HTMLBody
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('HTMLBody')]
        [string]$Html
    )
    process {
        if (-not $Html) { return '' }

        $md = $Html

        # Remove HTML comments (including MSO conditionals)
        $md = [regex]::Replace($md, '<!--.*?-->', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)

        # Remove <style> blocks entirely (MSO CSS bloat)
        $md = [regex]::Replace($md, '<style[^>]*>.*?</style>', '', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Remove <script> blocks
        $md = [regex]::Replace($md, '<script[^>]*>.*?</script>', '', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Remove <head> block
        $md = [regex]::Replace($md, '<head[^>]*>.*?</head>', '', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert <br> and <br/> to newlines
        $md = [regex]::Replace($md, '<br\s*/?>', "`n", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert <hr> to markdown
        $md = [regex]::Replace($md, '<hr\s*/?>', "`n---`n", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert headers h1-h6
        for ($i = 6; $i -ge 1; $i--) {
            $prefix = '#' * $i
            $md = [regex]::Replace($md, "<h$i[^>]*>(.*?)</h$i>", "`n$prefix `$1`n", [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        }

        # Convert <b> and <strong> to **bold**
        $md = [regex]::Replace($md, '<(?:b|strong)[^>]*>(.*?)</(?:b|strong)>', '**$1**', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert <i> and <em> to _italic_
        $md = [regex]::Replace($md, '<(?:i|em)[^>]*>(.*?)</(?:i|em)>', '_$1_', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert <u> to markdown (no native underline, use emphasis)
        $md = [regex]::Replace($md, '<u[^>]*>(.*?)</u>', '_$1_', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert <a href="url">text</a> to [text](url)
        $md = [regex]::Replace($md, '<a\s[^>]*href\s*=\s*"([^"]*)"[^>]*>(.*?)</a>', '[$2]($1)', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $md = [regex]::Replace($md, "<a\s[^>]*href\s*=\s*'([^']*)'[^>]*>(.*?)</a>", '[$2]($1)', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert images to markdown
        $md = [regex]::Replace($md, '<img\s[^>]*src\s*=\s*"([^"]*)"[^>]*alt\s*=\s*"([^"]*)"[^>]*/?\s*>', '![$2]($1)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $md = [regex]::Replace($md, '<img\s[^>]*src\s*=\s*"([^"]*)"[^>]*/?\s*>', '![image]($1)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert unordered lists
        $md = [regex]::Replace($md, '<li[^>]*>(.*?)</li>', "- `$1`n", [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $md = [regex]::Replace($md, '</?[uo]l[^>]*>', "`n", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # --- Table conversion ---
        $md = [regex]::Replace($md, '<table[^>]*>(.*?)</table>', {
            param($tableMatch)
            $tableHtml = $tableMatch.Groups[1].Value

            $rows = [regex]::Matches($tableHtml, '<tr[^>]*>(.*?)</tr>', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($rows.Count -eq 0) { return '' }

            $mdRows = @()
            foreach ($row in $rows) {
                $cells = [regex]::Matches($row.Groups[1].Value, '<t[hd][^>]*>(.*?)</t[hd]>', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                $cellTexts = @()
                foreach ($cell in $cells) {
                    $cellText = $cell.Groups[1].Value -replace '<[^>]+>', '' -replace '&nbsp;', ' '
                    $cellText = $cellText.Trim()
                    $cellTexts += $cellText
                }
                if ($cellTexts.Count -gt 0) {
                    $mdRows += '| ' + ($cellTexts -join ' | ') + ' |'
                }
            }

            if ($mdRows.Count -eq 0) { return '' }

            # Insert separator after first row (header)
            $colCount = ($mdRows[0] -split '\|').Count - 2  # minus leading/trailing empty
            $sep = '| ' + (('---') * [Math]::Max(1, $colCount) -join ' | ') + ' |'
            $result = "`n" + $mdRows[0] + "`n" + $sep
            for ($idx = 1; $idx -lt $mdRows.Count; $idx++) {
                $result += "`n" + $mdRows[$idx]
            }
            $result + "`n"
        }, [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Convert block elements to newlines
        $md = [regex]::Replace($md, '</?(?:p|div|tr|blockquote)[^>]*>', "`n", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        # Decode common HTML entities
        $md = $md -replace '&nbsp;', ' '
        $md = $md -replace '&amp;', '&'
        $md = $md -replace '&lt;', '<'
        $md = $md -replace '&gt;', '>'
        $md = $md -replace '&quot;', '"'
        $md = $md -replace '&#39;', "'"
        $md = $md -replace '&ndash;', '–'
        $md = $md -replace '&mdash;', '—'
        $md = $md -replace '&bull;', '•'
        $md = $md -replace '&#\d+;', ''

        # Strip all remaining HTML tags
        $md = [regex]::Replace($md, '<[^>]+>', '')

        # Clean up whitespace: collapse multiple blank lines to max 2
        $md = [regex]::Replace($md, '(\r?\n\s*){3,}', "`n`n")

        # Trim leading/trailing whitespace
        $md = $md.Trim()

        $md
    }
}

function Save-OutlookMail {
    <#
    .SYNOPSIS
    Save an Outlook email to disk in various formats.
    .EXAMPLE
    Save-OutlookMail -EntryID $id -Format MSG -DestinationPath C:\Temp
    .EXAMPLE
    Save-OutlookMail -EntryID $id -Format Markdown -DestinationPath C:\Temp
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$EntryID,

        [ValidateSet('MSG', 'HTML', 'TXT', 'Markdown')]
        [string]$Format = 'MSG',

        [string]$DestinationPath = (Join-Path $env:USERPROFILE 'Downloads'),

        [string]$FileName
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $item = $script:Namespace.GetItemFromID($EntryID)
    if (-not $item) { throw "Mail item not found." }

    if (-not (Test-Path $DestinationPath)) {
        New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
    }

    # Build safe filename from subject
    $safeName = if ($FileName) { $FileName } else {
        $s = $item.Subject -replace '[\\/:*?"<>|]', '_'
        if ($s.Length -gt 80) { $s = $s.Substring(0, 80) }
        $s
    }

    switch ($Format) {
        'MSG' {
            $path = Join-Path $DestinationPath "$safeName.msg"
            $item.SaveAs($path, 3)  # olMSG = 3
        }
        'HTML' {
            $path = Join-Path $DestinationPath "$safeName.html"
            $item.SaveAs($path, 5)  # olHTML = 5
        }
        'TXT' {
            $path = Join-Path $DestinationPath "$safeName.txt"
            $item.SaveAs($path, 0)  # olTXT = 0
        }
        'Markdown' {
            $path = Join-Path $DestinationPath "$safeName.md"
            $header = @"
# $($item.Subject)

**From:** $($item.SenderName) <$($item.SenderEmailAddress)>
**To:** $($item.To)
$(if ($item.CC) { "**CC:** $($item.CC)`n" })**Date:** $($item.ReceivedTime)
$(if ($item.Attachments.Count -gt 0) { "**Attachments:** $($item.Attachments | ForEach-Object { $_.FileName } | Join-String -Separator ', ')`n" })
---

"@
            $body = ConvertTo-EmailMarkdown -Html $item.HTMLBody
            Set-Content -Path $path -Value ($header + $body) -Encoding UTF8
        }
    }

    Write-Host "Saved: $path" -ForegroundColor Green
    [PSCustomObject]@{
        Path     = $path
        Format   = $Format
        Subject  = $item.Subject
        Size     = (Get-Item $path).Length
    }
}

Export-ModuleMember -Function Connect-Outlook, Disconnect-Outlook, Get-OutlookProfile,
    Get-OutlookFolder, Get-OutlookMail, Read-OutlookMail, Save-OutlookAttachment,
    Send-OutlookReply, Send-OutlookMail, ConvertTo-EmailMarkdown, Save-OutlookMail,
    Get-OutlookCalendar, New-OutlookAppointment
