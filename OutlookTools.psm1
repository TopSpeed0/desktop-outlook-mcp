$script:Outlook = $null
$script:Namespace = $null
$script:ModuleRoot = $PSScriptRoot
$script:DefaultMailbox = $null

$cfgPath = Join-Path $PSScriptRoot 'outlook-config.json'
if (Test-Path $cfgPath) {
    $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
    $script:DefaultMailbox = $cfg.mailbox
}

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

    $items = $folder.Items
    $items.Sort('[ReceivedTime]', $true)

    $collected = 0
    foreach ($item in $items) {
        if ($collected -ge $Count) { break }
        if ($item.Class -ne 43) { continue } # 43 = olMail
        if ($UnreadOnly -and $item.UnRead -eq $false) { continue }
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
            BodyPreview  = ($item.Body -replace '\r?\n', ' ').Substring(0, [Math]::Min(200, $item.Body.Length))
        }
        $collected++
    }
}

function Read-OutlookMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$EntryID
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $item = $script:Namespace.GetItemFromID($EntryID)
    if (-not $item) { throw "Mail item not found." }

    [PSCustomObject]@{
        EntryID      = $item.EntryID
        Subject      = $item.Subject
        From         = $item.SenderName
        FromEmail    = $item.SenderEmailAddress
        To           = $item.To
        CC           = $item.CC
        ReceivedTime = $item.ReceivedTime
        UnRead       = $item.UnRead
        Body         = $item.Body
        HTMLBody     = $item.HTMLBody
        Attachments  = @($item.Attachments | ForEach-Object { $_.FileName })
    }
}

function Send-OutlookReply {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$EntryID,

        [Parameter(Mandatory)]
        [string]$Body,

        [switch]$ReplyAll,

        [switch]$Send
    )
    if (-not $script:Namespace) { Connect-Outlook | Out-Null }
    $item = $script:Namespace.GetItemFromID($EntryID)
    if (-not $item) { throw "Mail item not found." }

    $reply = if ($ReplyAll) { $item.ReplyAll() } else { $item.Reply() }
    $reply.HTMLBody = $Body + $reply.HTMLBody

    if ($Send) {
        if ($PSCmdlet.ShouldProcess("Reply to '$($item.Subject)' from $($item.SenderName)", "Send")) {
            $reply.Send()
            Write-Host "Reply sent to '$($item.Subject)'." -ForegroundColor Green
            return [PSCustomObject]@{ Status = 'Sent'; Subject = $item.Subject; To = $item.SenderName }
        }
    }
    else {
        $reply.Display()
        Write-Host "Reply draft opened for '$($item.Subject)'." -ForegroundColor Cyan
        return [PSCustomObject]@{ Status = 'Draft'; Subject = $item.Subject; To = $item.SenderName }
    }
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

        [string[]]$Attachments,

        [switch]$HTML,

        [switch]$Send
    )
    if (-not $script:Outlook) { Connect-Outlook | Out-Null }

    $mail = $script:Outlook.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    if ($CC) { $mail.CC = $CC }
    if ($HTML) { $mail.HTMLBody = $Body } else { $mail.Body = $Body }

    foreach ($att in $Attachments) {
        if (Test-Path $att) { $mail.Attachments.Add($att) | Out-Null }
        else { Write-Warning "Attachment not found: $att" }
    }

    if ($Send) {
        if ($PSCmdlet.ShouldProcess("Send mail '$Subject' to $To", "Send")) {
            $mail.Send()
            Write-Host "Mail sent: '$Subject' to $To" -ForegroundColor Green
            return [PSCustomObject]@{ Status = 'Sent'; Subject = $Subject; To = $To }
        }
    }
    else {
        $mail.Display()
        Write-Host "Draft opened: '$Subject' to $To" -ForegroundColor Cyan
        return [PSCustomObject]@{ Status = 'Draft'; Subject = $Subject; To = $To }
    }
}

Export-ModuleMember -Function Connect-Outlook, Disconnect-Outlook, Get-OutlookProfile,
    Get-OutlookFolder, Get-OutlookMail, Read-OutlookMail, Send-OutlookReply, Send-OutlookMail
