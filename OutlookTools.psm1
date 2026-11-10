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

# Put the new text inside the reply's <body>, ahead of Outlook's signature and the quoted
# thread. Prepending to HTMLBody would land it before <html>, outside the document.
function Join-ReplyHtml {
    param([string]$Body, [string]$ReplyHtml)
    $m = [regex]::Match($ReplyHtml, '(?is)<body\b[^>]*>')
    if ($m.Success) { return $ReplyHtml.Insert($m.Index + $m.Length, $Body) }
    return $Body + $ReplyHtml
}

# True when Outlook already put the user's signature on the item (classic _MailAutoSig
# bookmark, or the Signature div newer builds use).
function Test-OutlookSignature {
    param([string]$Html)
    return [bool]($Html -match '(?i)_MailAutoSig|<div\b[^>]*\bid="?Signature\b')
}

# Remove a trailing "Thanks, / Regards, / Best, <Name>" that an AI wrote at the end of a
# body. Only a sign-off at the start of a line or block counts, so "Confirmed, thanks."
# stays. Returns the body unchanged if nothing matched or nothing else would be left.
function Remove-TrailingSignOff {
    param([string]$Html)
    $word = '(?:thanks|thank you|many thanks|thx|regards|best regards|kind regards|warm regards|best|cheers|br)'
    $gap  = '(?:\s|&nbsp;|<br\s*/?>|</?(?:p|div|span|b|strong|i|em|font)\b[^>]*>)*'
    $name = '(?-i:\p{Lu}[\p{L}.''-]*(?:\s+\p{Lu}[\p{L}.''-]*)?)'
    # Either "Thanks," + optional name, or "Thanks" alone on its line.
    $rx   = "(?is)(?<=^|>)(?:\s|&nbsp;)*$word(?:\s*[,!.]$gap(?:$name$gap)?|$gap)$"
    $m = [regex]::Match($Html, $rx)
    if (-not $m.Success) { return $Html }

    $kept = $Html.Substring(0, $m.Index)
    $kept = [regex]::Replace($kept, '(?is)(?:\s|&nbsp;|<br\s*/?>|<(?:p|div)\b[^>]*>)+$', '')
    $text = [regex]::Replace($kept, '<[^>]+>|&nbsp;|\s', '')
    if (-not $text) { return $Html }

    # Re-close any <p>/<div> whose closing tag went with the sign-off.
    foreach ($tag in 'p', 'div') {
        $open  = [regex]::Matches($kept, "(?i)<$tag\b").Count
        $close = [regex]::Matches($kept, "(?i)</$tag>").Count
        if ($open -gt $close) { $kept += "</$tag>" * ($open - $close) }
    }
    return $kept
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
            Date         = $item.ReceivedTime
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
        Date         = $item.ReceivedTime
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

function Set-OutlookMailRead {
    <#
    .SYNOPSIS
        Mark emails as read (or unread with -Unread).
    .EXAMPLE
        Get-OutlookMail -UnreadOnly -From 'Jira' | Set-OutlookMailRead
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$EntryID,

        [switch]$Unread
    )
    begin {
        if (-not $script:Namespace) { Connect-Outlook | Out-Null }
        $state = if ($Unread) { 'unread' } else { 'read' }
    }
    process {
        $item = $script:Namespace.GetItemFromID($EntryID)
        if (-not $item) { throw "Mail item not found." }

        if ($PSCmdlet.ShouldProcess("$($item.SenderName) | $($item.Subject)", "Mark as $state")) {
            if ($item.UnRead -ne [bool]$Unread) {
                $item.UnRead = [bool]$Unread
                $item.Save()
            }
        }

        [pscustomobject]@{
            Subject = $item.Subject
            From    = $item.SenderName
            UnRead  = $item.UnRead
        }
    }
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

        [string[]]$Attachments,

        [switch]$Send,

        [switch]$Unencrypted,

        [switch]$KeepLabel,

        [switch]$KeepSignOff
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $item = $script:Namespace.GetItemFromID($EntryID)
    if (-not $item) { throw "Mail item not found." }

    $reply = if ($ReplyAll) { $item.ReplyAll() } else { $item.Reply() }
    $replyHtml = $reply.HTMLBody

    # Outlook already added the user's signature, so a written sign-off would show twice.
    $signOffRemoved = $false
    if (-not $KeepSignOff -and (Test-OutlookSignature -Html $replyHtml)) {
        $trimmed = Remove-TrailingSignOff -Html $Body
        if ($trimmed -ne $Body) {
            $Body = $trimmed
            $signOffRemoved = $true
            Write-Host "Removed the sign-off from the body - Outlook adds the signature. Use -KeepSignOff to override." -ForegroundColor Yellow
        }
    }
    $reply.HTMLBody = Join-ReplyHtml -Body $Body -ReplyHtml $replyHtml

    foreach ($att in $Attachments) {
        if (Test-Path $att) { $reply.Attachments.Add($att) | Out-Null }
        else { Write-Warning "Attachment not found: $att" }
    }

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
            return [PSCustomObject]@{ Status = 'Sent'; Subject = $item.Subject; To = $item.SenderName; Recipients = $addresses; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients; SignOffRemoved = $signOffRemoved }
        }
        # -WhatIf / declined at the -Confirm prompt: discard, change nothing.
        Remove-UncommittedOutlookItem -MailItem $reply
        return [PSCustomObject]@{ Status = 'NotSent'; Subject = $item.Subject; To = $item.SenderName; Recipients = $addresses; Encryption = "Would be: $($decision.Encryption)"; ExternalRecipients = $decision.ExternalRecipients; SignOffRemoved = $signOffRemoved }
    }

    # -WhatIf without -Send: show what would happen, but open no draft.
    if ($WhatIfPreference) {
        Write-Host "What if: Opening reply draft for '$($item.Subject)'." -ForegroundColor Cyan
        Remove-UncommittedOutlookItem -MailItem $reply
        return [PSCustomObject]@{ Status = 'NotSent'; Subject = $item.Subject; To = $item.SenderName; Recipients = $addresses; Encryption = "Would be: $($decision.Encryption)"; ExternalRecipients = $decision.ExternalRecipients; SignOffRemoved = $signOffRemoved }
    }

    Invoke-OutlookEncryptionDecision -MailItem $reply -Decision $decision
    $reply.Display()
    Write-Host "Reply draft opened for '$($item.Subject)'." -ForegroundColor Cyan
    [PSCustomObject]@{ Status = 'Draft'; Subject = $item.Subject; To = $item.SenderName; Recipients = $addresses; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients; SignOffRemoved = $signOffRemoved }
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

        [switch]$KeepLabel,

        [switch]$KeepSignOff,

        [switch]$NoSignature
    )
    if (-not $script:Outlook) { Connect-Outlook | Out-Null }

    $mail = $script:Outlook.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    if ($CC) { $mail.CC = $CC }
    if ($BCC) { $mail.BCC = $BCC }

    # Match recipients to the address book now, as the Outlook GUI does while typing.
    $unresolved = @()
    try {
        $null = $mail.Recipients.ResolveAll()
        foreach ($r in $mail.Recipients) { if (-not $r.Resolved) { $unresolved += $r.Name } }
    }
    catch { Write-Verbose "Could not resolve recipients: $($_.Exception.Message)" }
    if ($unresolved) {
        Write-Warning "Outlook could not resolve: $($unresolved -join '; ')"
        if ($Send -and -not $WhatIfPreference) {
            Remove-UncommittedOutlookItem -MailItem $mail
            throw "Not sent - unresolved recipient(s): $($unresolved -join '; ')"
        }
    }

    # Outlook inserts the new-mail signature only when the inspector is created.
    $signatureAdded = $false
    if (-not $NoSignature) {
        try { $null = $mail.GetInspector } catch { Write-Verbose "Could not create inspector: $($_.Exception.Message)" }
        $signatureAdded = Test-OutlookSignature -Html $mail.HTMLBody
    }

    $htmlBody = if ($HTML) { $Body } else {
        '<p style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' +
        ([System.Net.WebUtility]::HtmlEncode($Body) -replace "\r?\n", '<br>') + '</p>'
    }

    $signOffRemoved = $false
    if ($signatureAdded) {
        if (-not $KeepSignOff) {
            $trimmed = Remove-TrailingSignOff -Html $htmlBody
            if ($trimmed -ne $htmlBody) {
                $htmlBody = $trimmed
                $signOffRemoved = $true
                Write-Host "Removed the sign-off from the body - Outlook adds the signature. Use -KeepSignOff to override." -ForegroundColor Yellow
            }
        }
        $mail.HTMLBody = Join-ReplyHtml -Body $htmlBody -ReplyHtml $mail.HTMLBody
    }
    elseif ($HTML) { $mail.HTMLBody = $Body }
    else { $mail.Body = $Body }

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
            return [PSCustomObject]@{ Status = 'Sent'; Subject = $Subject; To = $To; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients; SignatureAdded = $signatureAdded; SignOffRemoved = $signOffRemoved; Unresolved = $unresolved }
        }
        # -WhatIf / declined at the -Confirm prompt: discard, change nothing.
        Remove-UncommittedOutlookItem -MailItem $mail
        return [PSCustomObject]@{ Status = 'NotSent'; Subject = $Subject; To = $To; Encryption = "Would be: $($decision.Encryption)"; ExternalRecipients = $decision.ExternalRecipients; SignatureAdded = $signatureAdded; SignOffRemoved = $signOffRemoved; Unresolved = $unresolved }
    }

    # -WhatIf without -Send: show what would happen, but open no draft.
    if ($WhatIfPreference) {
        Write-Host "What if: Opening draft '$Subject' to $To." -ForegroundColor Cyan
        Remove-UncommittedOutlookItem -MailItem $mail
        return [PSCustomObject]@{ Status = 'NotSent'; Subject = $Subject; To = $To; Encryption = "Would be: $($decision.Encryption)"; ExternalRecipients = $decision.ExternalRecipients; SignatureAdded = $signatureAdded; SignOffRemoved = $signOffRemoved; Unresolved = $unresolved }
    }

    Invoke-OutlookEncryptionDecision -MailItem $mail -Decision $decision
    $mail.Display()
    Write-Host "Draft opened: '$Subject' to $To" -ForegroundColor Cyan
    [PSCustomObject]@{ Status = 'Draft'; Subject = $Subject; To = $To; Encryption = $decision.Encryption; ExternalRecipients = $decision.ExternalRecipients; SignatureAdded = $signatureAdded; SignOffRemoved = $signOffRemoved; Unresolved = $unresolved }
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

function Set-OutlookMeetingResponse {
    <#
    .SYNOPSIS
        Accept / Tentative / Decline a meeting invitation.
    .EXAMPLE
        Get-OutlookCalendar -Days 14 | Where-Object Organizer -match 'Noga' |
            Set-OutlookMeetingResponse -Response Accept -Send
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]$EntryID,

        [Parameter(Mandatory)]
        [ValidateSet('Accept', 'Tentative', 'Decline')]
        [string]$Response,

        [string]$Comment,

        # Occurrence start; used to target one occurrence of a recurring series
        [Parameter(ValueFromPipelineByPropertyName)]
        [Nullable[datetime]]$Start,

        [switch]$Send,
        [switch]$NoResponse
    )
    begin {
        if (-not $script:Namespace) { Connect-Outlook | Out-Null }
        $map = @{ Accept = 3; Tentative = 2; Decline = 4 }  # OlMeetingResponse
    }
    process {
        $item = $script:Namespace.GetItemFromID($EntryID)

        if ($item.MessageClass -like 'IPM.Schedule.Meeting.Request*') {
            $item = $item.GetAssociatedAppointment($true)
        }
        if (-not $item -or $item.MessageClass -notlike 'IPM.Appointment*') {
            throw "EntryID is not a meeting/appointment: $EntryID"
        }

        if ($item.IsRecurring -and $Start) {
            $item = $item.GetRecurrencePattern().GetOccurrence([datetime]$Start)
        }

        # MeetingStatus: 0 = olNonMeeting, 1 = olMeeting (organizer), 3 = olMeetingReceived
        if ($item.MeetingStatus -eq 0) {
            throw "'$($item.Subject)' is not a meeting."
        }
        if ($item.MeetingStatus -eq 1) {
            throw "You are the organizer of '$($item.Subject)' - nothing to respond to."
        }

        $target = "$($item.Subject) @ $($item.Start) (organizer: $($item.Organizer))"
        if (-not $PSCmdlet.ShouldProcess($target, "$Response meeting")) {
            return [pscustomobject]@{
                Subject = $item.Subject; Start = $item.Start; Organizer = $item.Organizer
                Response = $Response; Status = 'NotSent'
            }
        }

        $reply = $item.Respond($map[$Response], $true, $false)

        $status = 'CalendarOnly'
        if ($reply -and -not $NoResponse) {
            if ($Comment) { $reply.Body = $Comment }
            if ($Send) { $reply.Send(); $status = 'Sent' }
            else       { $reply.Display(); $status = 'DraftOpened' }
        }

        [pscustomobject]@{
            Subject   = $item.Subject
            Start     = $item.Start
            Organizer = $item.Organizer
            Response  = $Response
            Status    = $status
        }
    }
}

Set-Alias -Name Respond-OutlookMeeting -Value Set-OutlookMeetingResponse

# ---------------------------------------------------------------------------
# Review workbook: -reviewCSV / -SendCSV / -SendCSVRAW
# ---------------------------------------------------------------------------

$script:ReviewFolder = Join-Path $env:USERPROFILE 'Downloads\OutlookReview'
$script:ReviewDecisions = 'Reply', 'Ignore', 'Done', 'False Alarm'
# Rows with one of these decisions are never drafted and not shown again by the next review.
$script:ReviewClosedDecision = '^(ignore|done|false alarm)$'

# A row is pending when it has Reply text, is not closed by its Decision, and was not drafted.
function Test-OutlookReviewPending {
    param($Row)
    "$($Row.Reply)".Trim() -and
        "$($Row.Decision)".Trim() -notmatch $script:ReviewClosedDecision -and
        "$($Row.Status)" -notmatch '^Drafted'
}

# EntryIDs the user already handled in earlier workbooks (closed decision or drafted reply).
function Get-OutlookReviewHandledIds {
    $ids = @{}
    foreach ($f in Get-ChildItem $script:ReviewFolder -Filter 'OutlookReview-*.xlsx' -ErrorAction SilentlyContinue) {
        try { $rows = Import-Excel -Path $f.FullName -WorksheetName 'Review' -ErrorAction Stop }
        catch { Write-Warning "Could not read $($f.Name) (open in Excel?) - its decisions are not applied."; continue }
        foreach ($r in $rows) {
            if (-not $r.EntryID) { continue }
            $why = if ("$($r.Decision)".Trim() -match $script:ReviewClosedDecision) { "$($r.Decision)".Trim() }
                   elseif ("$($r.Status)" -match '^Drafted') { 'Drafted' }
            if ($why) { $ids[$r.EntryID] = "$why ($($f.Name))" }
        }
    }
    $ids
}

function Import-ImportExcelModule {
    if (-not (Get-Module ImportExcel)) {
        try { Import-Module ImportExcel -ErrorAction Stop }
        catch { throw "The ImportExcel module is required: Install-Module ImportExcel -Scope CurrentUser" }
    }
}

# Newest review workbook, used when -Path is not given.
function Resolve-OutlookReviewPath {
    param([string]$Path)
    if ($Path) {
        if (-not (Test-Path $Path)) { throw "Review file not found: $Path" }
        return (Resolve-Path $Path).Path
    }
    $latest = Get-ChildItem $script:ReviewFolder -Filter 'OutlookReview-*.xlsx' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { throw "No review file found in $script:ReviewFolder. Run the review first or pass -Path." }
    $latest.FullName
}

function Assert-OutlookReviewWritable {
    param([string]$Path)
    try { $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None'); $fs.Close() }
    catch { throw "The review file is open or locked - close it in Excel and run again: $Path" }
}

# The newest message only: cut the quoted thread so the agent reads less.
function Get-LatestMessageText {
    param([string]$Text, [int]$MaxLength)
    if (-not $Text) { return '' }
    $m = [regex]::Match($Text, '(?im)^\s*(From:\s.+|-{3,}\s*Original Message\s*-{3,}|On .{5,120} wrote:)\s*$')
    if ($m.Success -and $m.Index -gt 0) { $Text = $Text.Substring(0, $m.Index) }
    $Text = ($Text -replace '(\r?\n\s*){3,}', "`n`n").Trim()
    if ($MaxLength -gt 0 -and $Text.Length -gt $MaxLength) { $Text = $Text.Substring(0, $MaxLength) + ' [...]' }
    $Text
}

function Get-OutlookReviewCandidates {
    <#
    .SYNOPSIS
    Collect mail addressed to the user (To or CC) for the AI to triage into a review workbook.
    .DESCRIPTION
    Default: unread Inbox mail. -Days N scans read and unread mail from the last N days instead.
    Mail sent by the user and mail where the user is only on a distribution list are skipped.
    Mail already marked Ignore / Done / False Alarm, or already drafted, in an earlier review
    workbook is skipped too (-IncludeHandled shows it again).
    Reading items here does not mark them read.
    #>
    [CmdletBinding()]
    param(
        [string]$Mailbox = $script:DefaultMailbox,
        [string]$FolderPath = 'Inbox',
        [int]$Days = 0,
        [int]$MaxItems = 200,
        [int]$MaxBodyLength = 1500,
        [switch]$IncludeHandled
    )
    if (-not $Mailbox) { throw "No mailbox specified. Pass -Mailbox or set mailbox in outlook-config.json." }
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $handled = @{}
    if (-not $IncludeHandled) { Import-ImportExcelModule; $handled = Get-OutlookReviewHandledIds }
    $skipped = 0

    $folder = $script:Namespace.Session.Folders.Item($Mailbox)
    foreach ($part in ($FolderPath -split '\\' | Where-Object { $_ })) {
        $found = $null
        foreach ($f in $folder.Folders) { if ($f.Name -eq $part) { $found = $f; break } }
        if (-not $found) { throw "Folder '$part' not found." }
        $folder = $found
    }

    $filter = if ($Days -gt 0) {
        "[ReceivedTime] >= '$(Format-OutlookDate -Date (Get-Date).AddDays(-$Days))'"
    } else {
        '@SQL=' + (New-DaslEqualsClause -Property $script:Dasl.Read -Value 0)
    }
    $items = $folder.Items.Restrict($filter)
    try { $items.Sort('[ReceivedTime]', $true) } catch {}

    $me = $Mailbox.ToLowerInvariant()
    $n = 0
    foreach ($item in $items) {
        if ($n -ge $MaxItems) { Write-Warning "Stopped at -MaxItems $MaxItems."; break }
        if ($item.Class -ne 43) { continue }
        if ($handled.ContainsKey($item.EntryID)) {
            Write-Verbose "Skipped, already handled: $($handled[$item.EntryID]) - $($item.Subject)"
            $skipped++; continue
        }
        if (-not $item.Subject -and $item.EntryID) {
            try { $item = $script:Namespace.GetItemFromID($item.EntryID) } catch {}
        }

        $sender = $null
        try { $sender = $item.Sender.GetExchangeUser().PrimarySmtpAddress } catch {}
        if (-not $sender) { $sender = $item.SenderEmailAddress }
        if ($sender -and $sender.ToLowerInvariant() -eq $me) { continue }

        $role = $null
        foreach ($r in $item.Recipients) {
            $smtp = Resolve-RecipientSmtp -Recipient $r
            if ($smtp -and $smtp.ToLowerInvariant() -eq $me) {
                $t = if ($r.Type -eq 1) { 'To' } elseif ($r.Type -eq 2) { 'CC' } else { $null }
                if ($t -eq 'To') { $role = 'To'; break }
                if ($t) { $role = $t }
            }
        }
        if (-not $role) { continue }

        $n++
        [PSCustomObject]@{
            EntryID        = $item.EntryID
            Date           = $item.ReceivedTime
            From           = $item.SenderName
            FromEmail      = $sender
            Subject        = $item.Subject
            Role           = $role
            ToCount        = ($item.To -split ';' | Where-Object { $_.Trim() }).Count
            CcCount        = ($item.CC -split ';' | Where-Object { $_.Trim() }).Count
            Importance     = @('Low', 'Normal', 'High')[$item.Importance]
            HasAttachments = $item.Attachments.Count -gt 0
            UnRead         = $item.UnRead
            Body           = Get-LatestMessageText -Text $item.Body -MaxLength $MaxBodyLength
        }
    }
    if ($skipped) { Write-Host "Skipped $skipped mail(s) already handled in earlier review workbooks (-IncludeHandled shows them)." -ForegroundColor Cyan }
}

function Export-OutlookReview {
    <#
    .SYNOPSIS
    Write the triaged mail to a formatted review workbook (.xlsx).
    .DESCRIPTION
    Each input row needs EntryID, Date, From, Subject, Role, plus the AI's Priority
    (High/Medium/Low), Category (Action/Info), Summary and ActionNeeded. Reply, ReplyAll
    and Status start empty for the user to fill.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]]$InputObject,
        [string]$Path,
        [switch]$Show
    )
    begin { $rows = [System.Collections.Generic.List[object]]::new() }
    process { foreach ($o in $InputObject) { $rows.Add($o) } }
    end {
        Import-ImportExcelModule
        if (-not $Path) {
            New-Item -ItemType Directory -Path $script:ReviewFolder -Force | Out-Null
            $Path = Join-Path $script:ReviewFolder ("OutlookReview-{0:yyyyMMdd-HHmm}.xlsx" -f (Get-Date))
        }
        if (Test-Path $Path) { Assert-OutlookReviewWritable -Path $Path; Remove-Item $Path -Force }

        $rank = @{ High = 0; Medium = 1; Low = 2 }
        $i = 0
        $data = $rows | Sort-Object @{ e = { if ($rank.ContainsKey("$($_.Priority)")) { $rank["$($_.Priority)"] } else { 3 } } },
                                    @{ e = { $_.Date }; Descending = $true } | ForEach-Object {
            $i++
            [PSCustomObject][ordered]@{
                No = $i; Priority = $_.Priority; Category = $_.Category
                Date = if ($_.Date -is [double]) { [datetime]::FromOADate($_.Date) } elseif ($_.Date) { [datetime]$_.Date } else { $null }
                From = $_.From; Subject = $_.Subject; Role = $_.Role
                Summary = $_.Summary; ActionNeeded = $_.ActionNeeded
                Decision = "$($_.Decision)"; Reply = "$($_.Reply)"
                ReplyAll = if ($_.ReplyAll) { "$($_.ReplyAll)" } else { 'No' }
                Notes = "$($_.Notes)"; Status = "$($_.Status)"; EntryID = $_.EntryID
            }
        }
        if (-not $data) { Write-Warning "Nothing to export."; return }

        # A No, B Priority, C Category, D Date, E From, F Subject, G Role, H Summary, I ActionNeeded,
        # J Decision, K Reply, L ReplyAll, M Notes, N Status, O EntryID (hidden)
        $pkg = $data | Export-Excel -Path $Path -WorksheetName 'Review' -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
        $ws = $pkg.Workbook.Worksheets['Review']
        $last = $data.Count + 1
        $widths = @{ A = 5; B = 9; C = 9; D = 16; E = 22; F = 40; G = 6; H = 60; I = 40; J = 12; K = 60; L = 9; M = 35; N = 18; O = 12 }
        foreach ($c in $widths.Keys) { $ws.Column([int][char]$c - 64).Width = $widths[$c] }
        $ws.Cells["D2:D$last"].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm'
        $ws.Cells["A1:O$last"].Style.VerticalAlignment = 'Top'
        foreach ($c in 'F', 'H', 'I', 'K', 'M') { $ws.Cells["${c}2:${c}$last"].Style.WrapText = $true }
        $ws.Cells["A1:O1"].Style.Fill.PatternType = 'Solid'
        $ws.Cells["A1:O1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0x1F, 0x4E, 0x79))
        $ws.Cells["A1:O1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        foreach ($c in 'J', 'K', 'L', 'M') {
            $ws.Cells["${c}2:${c}$last"].Style.Fill.PatternType = 'Solid'
            $ws.Cells["${c}2:${c}$last"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0xFF, 0xF8, 0xE1))
        }
        $ws.Column(15).Hidden = $true

        $prio = @{ High = 'FFCDD2'; Medium = 'FFE0B2'; Low = 'E8F5E9' }
        foreach ($p in $prio.Keys) {
            Add-ConditionalFormatting -Worksheet $ws -Address "B2:B$last" -RuleType Equal -ConditionValue "`"$p`"" `
                -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml("#$($prio[$p])"))
        }
        # A closed decision greys out the row.
        Add-ConditionalFormatting -Worksheet $ws -Address "A2:O$last" -RuleType Expression `
            -ConditionValue '=OR($J2="Ignore",$J2="Done",$J2="False Alarm")' -ForegroundColor Gray `
            -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml('#EEEEEE'))
        $d = $ws.DataValidations.AddListValidation("J2:J$last")
        foreach ($x in $script:ReviewDecisions) { $d.Formula.Values.Add($x) }
        $v = $ws.DataValidations.AddListValidation("L2:L$last")
        $v.Formula.Values.Add('Yes'); $v.Formula.Values.Add('No')

        Close-ExcelPackage $pkg -Show:$Show
        Write-Host "Review workbook: $Path ($($data.Count) mails)" -ForegroundColor Green
        [PSCustomObject]@{ Path = $Path; Count = $data.Count }
    }
}

function Import-OutlookReview {
    <#
    .SYNOPSIS
    Read a review workbook. -Pending returns only rows with a Reply, no Drafted status, and a
    Decision that is empty or Reply (Ignore / Done / False Alarm rows are never pending).
    #>
    [CmdletBinding()]
    param([string]$Path, [switch]$Pending)
    Import-ImportExcelModule
    $Path = Resolve-OutlookReviewPath -Path $Path
    $rows = Import-Excel -Path $Path -WorksheetName 'Review'
    if ($Pending) { $rows = $rows | Where-Object { Test-OutlookReviewPending -Row $_ } }
    $rows | ForEach-Object { $_ | Add-Member -NotePropertyName ReviewPath -NotePropertyValue $Path -PassThru }
}

# Plain text to HTML with nothing changed but line breaks: blank line = new paragraph,
# single newline = <br>, runs of spaces kept.
function ConvertTo-OutlookReplyHtml {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Text)
    $paras = ($Text.Trim() -replace '\r\n?', "`n") -split '\n\s*\n' | ForEach-Object {
        $p = [System.Net.WebUtility]::HtmlEncode($_.Trim()) -replace '  ', '&nbsp; ' -replace '\n', '<br>'
        "<p style=`"margin:0 0 10px 0`">$p</p>"
    }
    '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#333">' + ($paras -join '') + '</div>'
}

function Send-OutlookReviewReplies {
    <#
    .SYNOPSIS
    Open a reply DRAFT for each review row that has a Reply and is not yet drafted. Never sends.
    .DESCRIPTION
    -Raw: the Reply text is used as written; only line breaks become HTML, and a written
    sign-off is kept. Otherwise pass -HtmlByEntryID with the AI-written HTML for each row.
    Each drafted row gets Status "Drafted <time>" so a second run skips it.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$Path,
        [switch]$Raw,
        [hashtable]$HtmlByEntryID
    )
    if (-not $Raw -and -not $HtmlByEntryID) { throw "Pass -Raw, or -HtmlByEntryID with the AI-written reply HTML." }
    Import-ImportExcelModule
    $Path = Resolve-OutlookReviewPath -Path $Path
    Assert-OutlookReviewWritable -Path $Path

    $rows = @(Import-OutlookReview -Path $Path -Pending)
    if (-not $rows) { Write-Host "No pending replies in $Path." -ForegroundColor Cyan; return }

    $done = @{}
    $results = foreach ($row in $rows) {
        $html = if ($Raw) { ConvertTo-OutlookReplyHtml -Text "$($row.Reply)" } else { $HtmlByEntryID[$row.EntryID] }
        if (-not $html) {
            Write-Warning "No reply HTML for row $($row.No) '$($row.Subject)' - skipped."
            [PSCustomObject]@{ No = $row.No; Subject = $row.Subject; Status = 'Skipped' }
            continue
        }
        if (-not $PSCmdlet.ShouldProcess("Row $($row.No) '$($row.Subject)'", 'Open reply draft')) {
            [PSCustomObject]@{ No = $row.No; Subject = $row.Subject; Status = 'NotDrafted' }
            continue
        }
        try {
            $r = Send-OutlookReply -EntryID $row.EntryID -Body $html -ReplyAll:("$($row.ReplyAll)" -match '^(y|yes|true)$') -KeepSignOff:$Raw
            $done[$row.EntryID] = "Drafted {0:yyyy-MM-dd HH:mm}" -f (Get-Date)
            [PSCustomObject]@{ No = $row.No; Subject = $row.Subject; Status = $r.Status; Recipients = ($r.Recipients -join '; ') }
        }
        catch {
            Write-Warning "Row $($row.No) '$($row.Subject)': $($_.Exception.Message)"
            [PSCustomObject]@{ No = $row.No; Subject = $row.Subject; Status = 'Failed' }
        }
    }

    if ($done.Count) {
        $pkg = Open-ExcelPackage -Path $Path
        $ws = $pkg.Workbook.Worksheets['Review']
        # Find columns by header so older workbooks (without Decision/Notes) still work.
        $col = @{}
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) { $col["$($ws.Cells[1, $c].Value)"] = $c }
        for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
            $id = $ws.Cells[$r, $col['EntryID']].Value
            if ($id -and $done.ContainsKey($id)) { $ws.Cells[$r, $col['Status']].Value = $done[$id] }
        }
        Close-ExcelPackage $pkg
    }
    $results
}

Export-ModuleMember -Function Connect-Outlook, Disconnect-Outlook, Get-OutlookProfile,
    Get-OutlookFolder, Get-OutlookMail, Read-OutlookMail, Set-OutlookMailRead, Save-OutlookAttachment,
    Send-OutlookReply, Send-OutlookMail, ConvertTo-EmailMarkdown, Save-OutlookMail,
    Get-OutlookCalendar, New-OutlookAppointment, Set-OutlookMeetingResponse,
    Get-OutlookReviewCandidates, Export-OutlookReview, Import-OutlookReview,
    ConvertTo-OutlookReplyHtml, Send-OutlookReviewReplies `
    -Alias Respond-OutlookMeeting
