# Description: This script restores .msg files from a folder structure to Outlook folders.
# Author: yitzhak bohadana.
# all rights reserved.
# Version: 1.1.1
# Date: 2021-07-01
# Usage: Run the script and select the source folder containing the .msg files.
#
# $DEBUG = $true

$click_yes = Get-ChildItem .\click_yes.exe
try {
Start-Process $click_yes -WindowStyle Hidden -ErrorAction SilentlyContinue
}
catch {
    Write-Warning "An error occurred: $_"
    # Handle errors here if necessary
}

# Main logic of your script (Restore-MSG functionality)
Write-Host "Running Restore-MSG functionality..."
new-item -path $env:USERPROFILE\msg_logs -type directory -force -ErrorAction SilentlyContinue | out-null
function get-dateNow {
    $logdate = get-date -Format 'dd_MM_yy.mm.ss'
    return $logdate
} 
$msg_logs = "$env:USERPROFILE\msg_logs\msg_$(get-dateNow).log"
$msg_logsTemp = "$env:USERPROFILE\msg_logs\msg_$(get-dateNow)_temp.log"
"TIME: $(get-dateNow) | INFO | Start Import for User: $($env:USERNAME)" | Out-File -FilePath $msg_logs -Append


# get user decision function
function Get-UserDecision ($MessageboxTitle, $Messageboxbody) {
    Add-Type -AssemblyName PresentationCore, PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::YesNo
    if ($null -eq $MessageboxTitle) { $MessageboxTitle = "Yes or No" }
    if ($null -eq $Messageboxbody) { $Messageboxbody = "Are you sure you want to do this task ?" }
    
    $MessageIcon = [System.Windows.MessageBoxImage]::Warning
    [System.Windows.MessageBox]::Show($Messageboxbody, $MessageboxTitle, $ButtonType, $messageicon)
    
}
# select folder using windows forms
Function Select-FolderDialog {
    param([string]$Description = "Select Folder", [string]$RootFolder = "Desktop")

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null     
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = $RootFolder
    $objForm.Description = $Description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK") {
        Return $objForm.SelectedPath
    }
    Else {
        Write-Error "Operation cancelled by user."
    }
}
# Select from List GUI
function Select-FromList { 
    param (
        $ArryList,
        $MessageBOX,
        $testBox
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
        
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $MessageBOX
    $form.Size = New-Object System.Drawing.Size(600, 700)
    $form.StartPosition = 'CenterScreen'
        
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(500, 50)
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)
        
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(500, 75)
    $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)
        
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(280, 20)
    $label.Text = $testBox
    $form.Controls.Add($label)
        
    $listBox = New-Object System.Windows.Forms.Listbox
    $listBox.Location = New-Object System.Drawing.Point(10, 40)
    $listBox.Size = New-Object System.Drawing.Size(475, 100)
        
    $listBox.SelectionMode = 'MultiExtended'
        
    foreach ($Item in $ArryList) {
        [void] $listBox.Items.Add($Item)
    }
        
    $listBox.Height = 590
    $form.Controls.Add($listBox)
    $form.Topmost = $true
        
    $result = $form.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
        
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $x = $listBox.SelectedItems
        return  $x
    }
}

# Define the source directory containing the folder tree and .msg files
function Get-OrCreateFolder {
    param (
        [string]$folderPath,
        [object]$parentOutlookFolder
    )
    
    $parts = $folderPath -split '\\'
    $currentFolder = $parentOutlookFolder
    
    foreach ($part in $parts) {
        $subFolder = $null
        foreach ($folder in $currentFolder.Folders) {
            if ($folder.Name -eq $part) {
                $subFolder = $folder
                break
            }
        }    
        if (-not $subFolder) {
            $subFolder = $currentFolder.Folders.Add($part)
        }
        $currentFolder = $subFolder
    }
    return $currentFolder
}

# Define the source directory containing the folder tree and .msg files
$sourceDir = Select-FolderDialog -Description "Select the folder containing the .msg files" -RootFolder "Desktop"
"TIME: $(get-dateNow) | INFO | MSG Folder: $sourceDir Selected for import to User: $($env:USERNAME)" | Out-File -FilePath $msg_logs -Append


# Connect to Outlook
$MyEmail = $null
try { 
    $outlook = New-Object -ComObject Outlook.Application 
}
catch {
    Write-Error "Outlook is not responding please restart the Program and restart outlook !"
    pause
    exit
}
# $MailProperties = $outlook.Session.folders | Select-Object name | ? { ($_.name -match "@") -and ($_.name -notmatch "Public") }
$OutlookRootFolderSelected = Select-FromList -ArryList ($outlook.Session.folders | Select-Object name )  -MessageBOX "Select Profile." -testBox "Select Profile from the list"

# $MyEmail = $MailProperties.name
$MyEmail = $OutlookRootFolderSelected.name
$namespace = $outlook.GetNamespace('MAPI')

# get user decision if the folder is nested under Inbox or not

$YesNoRoot = Get-UserDecision -MessageboxTitle "has the restore is One to One of Top of Information Store ?" -Messageboxbody "Select Yes if the root folder of the Mailbox:$MyEmail will be selected"
if ($YesNoRoot -eq 'yes') {
    $rootFolder = $namespace.Session.folders.Item("$MyEmail") # root folder
}
else {
    $YesNo = Get-UserDecision -MessageboxTitle "has the Restored folder like is nested under Inbox ?" -Messageboxbody "Select Yes if the folder is nested under Inbox, otherwise select No"

    if ($YesNo -eq 'yes') {
        # if the folder is nested under Inbox
        $MailProperties = $namespace.Session.folders.Item("$MyEmail").Folders.Item('Inbox').Folders | Select-Object FullFolderPath
        $MailPropertiesF = foreach ($MailPropertie in $MailProperties) { ($MailPropertie.FullFolderPath).Split('\')[4] } # more underline aproach due to diffrent powershell version
        # $MailPropertiesF = foreach ($MailPropertie in $MailProperties) { ($MailPropertie.FullFolderPath).Split('Inbox\')[1] }
        $FolderSelected = Select-FromList -ArryList $MailPropertiesF  -MessageBOX "Select the target folder" -testBox "Select the target folder"
        $rootFolder = $namespace.Session.folders.Item("$MyEmail").Folders.Item('Inbox').Folders.Item($FolderSelected) # nested under Inbox
    }
    if ($YesNo -eq 'no') {
        # if the folder is not nested under Inbox (root folder)
        $MailProperties = $namespace.Session.folders.Item("$MyEmail").Folders | Select-Object FullFolderPath
        $MailPropertiesF = foreach ($MailPropertie in $MailProperties) { ($MailPropertie.FullFolderPath).Split('\')[3] } # more underline aproach due to diffrent powershell version
        # $MailPropertiesF = foreach ($MailPropertie in $MailProperties) { ($MailPropertie.FullFolderPath).Split('com\')[1] }
        $FolderSelected = Select-FromList -ArryList $MailPropertiesF  -MessageBOX "Select the target folder" -testBox "Select the target folder"
        $rootFolder = $namespace.Session.folders.Item("$MyEmail").Folders.Item($FolderSelected)# nested under Inbox
    }
}

###      ###
### Main ###
###      ###

if ($YesNoRoot -eq 'no') {
    # imort the .msg of the root folder
    Get-ChildItem -Path $sourceDir | ? { $_.name -like "*.msg" } | ForEach-Object {
        # $rootFolder
        $ERRORlog = $null
        $fatalError = $null
        $WarrningLog = $null
        $msgPath = $_.FullName
        try {
            $mailItem = $namespace.OpenSharedItem($msgPath)
            $INFO_LoadMSG = "TIME: $(get-dateNow) | INFO | trying to resore Item: <-> $($mailItem.Subject) <-> from: $msgPath to $($rootFolder.FolderPath)"
            $INFO_LoadMSG | Out-File -FilePath $msg_logs -Append
            Write-host $INFO_LoadMSG -ForegroundColor Blue
        }
        catch {
            $fatalError = "TIME: $(get-dateNow) | FATAL | Failed to open File MSG_Path:<-> $msgPath <-> to Load mess $($_.Exception.Message)" 
            $fatalError | Out-File -FilePath $msg_logs -Append
            Write-Host $fatalError -ForegroundColor Red
            # pause
        }
        if (!$fatalError) {
            try {
                $mailItem.Move($rootFolder) | Out-Null
                $namespace.close
            }
            catch {
                if ($($_.Exception.Message) -match "items were copied" ) {
                    $WarrningLog = "TIME: $(get-dateNow) | Warrning | items:$($mailItem.Subject) ->> were copied instead of moved, copied to Folder:$($rootFolder.FolderPath) | From MSG_Path:<-> $msgPath <-> Error Messages: $($_.Exception.Message)"  
                    $WarrningLog | Out-File -FilePath $msg_logs -Append
                    Write-Host $WarrningLog -ForegroundColor Yellow
                }
                else {
                    $ERRORlog = "TIME: $(get-dateNow) | ERROR | Failed to move Mail:$($mailItem.Subject) | MSG_Path:<-> $msgPath <-> to Folder::$($rootFolder.FolderPath) ,Error Messages: $($_.Exception.Message)" 
                    $ERRORlog | Out-File -FilePath $msg_logs -Append 
                    Write-Host $ERRORlog -ForegroundColor Red
                    # "Fallback: ">> $msg_logs ; $mailItem >> $msg_logsTemp 2>> $msg_logs
                }
            }
            finally {
                if (!$ERRORlog -and !$WarrningLog) {
                    $INFO_LoadMSG = "TIME: $(get-dateNow) | Successfully Restored Item:$($mailItem.Subject)"
                    $INFO_LoadMSG | Out-File -FilePath $msg_logs -Append 
                    Write-Host $INFO_LoadMSG -ForegroundColor Green

                    # Correctly renaming the file by removing .msg and adding .bkp
                    $newName = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName) + ".bkp"
                    Rename-Item -Path $_.FullName -NewName $newName -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

# Recursively process the folder structure and import .msg files
Get-ChildItem -Path $sourceDir -Recurse -Directory | ForEach-Object {
    $relativePath = $_.FullName.Substring($sourceDir.Length).TrimStart('\')
    $outlookFolder = Get-OrCreateFolder -folderPath $relativePath -parentOutlookFolder $rootFolder

    # Import .msg files into the corresponding folder
    Get-ChildItem -Path $_.FullName | ? { $_.name -like "*.msg" } | ForEach-Object {
        $fatalError = $null
        $ERRORlog = $null
        $WarrningLog = $null
        $msgPath = $_.FullName
        try {
            $mailItem = $namespace.OpenSharedItem($msgPath)
            $INFO_LoadMSG = "TIME: $(get-dateNow) | INFO | trying to resore Item: <-> $($mailItem.Subject) <-> from: $msgPath to $($outlookFolder.FolderPath)"
            Write-host $INFO_LoadMSG -ForegroundColor Blue
            $INFO_LoadMSG | Out-File -FilePath $msg_logs -Append
        } 
        catch {
            $fatalError = "TIME: $(get-dateNow) | FATAL | Failed to open File MSG_Path:<-> $msgPath <-> to Load mess $($_.Exception.Message)" 
            $fatalError | Out-File -FilePath $msg_logs -Append
            Write-Host $fatalError -ForegroundColor Red
        }
        if (!$fatalError) {
            try {
                if ($outlookFolder.name -eq 'inbox' -and $outlookFolder.Parent.name -eq $MyEmail) {
                    $INBOX = $namespace.GetDefaultFolder(6)
                    try { $mailItem.Move($INBOX) } catch {
                        $mailItem >> $msg_logsTemp 2>> $msg_logs
                    }
                    $namespace.close
                }
                else {
                    $mailItem.Move($outlookFolder) | Out-Null
                    $namespace.close
                }
            }
            catch {
                if ($($_.Exception.Message) -match "items were copied" ) {
                    $WarrningLog = "TIME: $(get-dateNow) | Warrning | items:$($mailItem.Subject) ->> were copied instead of moved, copied to Folder:$($outlookFolder.FolderPath) | From MSG_Path:<-> $msgPath <-> Error Messages: $($_.Exception.Message)"  
                    $WarrningLog | Out-File -FilePath $msg_logs -Append
                    Write-Host $WarrningLog -ForegroundColor Yellow
                }
                else {
                    $ERRORlog = "TIME: $(get-dateNow) | ERROR | Failed to move Mail:$($mailItem.Subject) | MSG_Path:<-> $msgPath <-> to Folder::$($outlookFolder.FolderPath) ,Error Messages: $($_.Exception.Message)" 
                    $ERRORlog | Out-File -FilePath $msg_logs -Append 
                    Write-Host $ERRORlog -ForegroundColor Red
                    # "Fallback: ">> $msg_logs ; $mailItem >> $msg_logsTemp 2>> $msg_logs
                }
            }
            finally {
                if (!$ERRORlog -and !$WarrningLog) {
                    $INFO_LoadMSG = "TIME: $(get-dateNow) | Successfully Restored Item:$($mailItem.Subject)"
                    $INFO_LoadMSG | Out-File -FilePath $msg_logs -Append 
                    Write-Host $INFO_LoadMSG -ForegroundColor Green
                    
                    # Correctly renaming the file by removing .msg and adding .bkp
                    $newName = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName) + ".bkp"
                    Rename-Item -Path $_.FullName -NewName $newName -ErrorAction SilentlyContinue
                } 
            }
        }
    }
}

# open log file
explorer $msg_logs
remove-item $msg_logsTemp -Force -ErrorAction SilentlyContinue

Get-Process "click_yes" | Stop-Process -Force -ErrorAction SilentlyContinue