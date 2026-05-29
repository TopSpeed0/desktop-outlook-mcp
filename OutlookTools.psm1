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
    Send-OutlookReply, Send-OutlookMail, ConvertTo-EmailMarkdown, Save-OutlookMail
