# Ensure this script runs in PowerShell Core
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell Core (version 7 or higher). Exiting."
    exit 1
}
cd "C:\Users\ybohadana\OneDrive - COGNYTE\Documents\code\Cognyte\Microsoft\Outlook\Restore-MSG"
Invoke-ps2exe -inputFile .\Restore-MSG.ps1 -outputFile .\Restore-MSG.exe -title "Restore MSG" -iconFile .\Designer.ico -company Cognyte -description "Restore MSG base file to outlook" -version 1.1.1 -copyright Cognyte