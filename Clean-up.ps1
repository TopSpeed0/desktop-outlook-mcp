#if DEBUG delete the Restore1 and Restore2 folders as they are just test for the script and not needed anymore
if ($DEBUG) {
    try {
        $tempFolder = $namespace.Session.Folders.Item("$MyEmail").Folders | Where-Object { $_.Name -eq "Restore1" }
        if ($tempFolder) {
            $tempFolder.Delete()
            "TIME: $(get-dateNow) | INFO | Deleted Restore1 folder after successful import." | Out-File -FilePath $msg_logs -Append
        }
    }
    catch {
        "TIME: $(get-dateNow)  | ERROR |: Failed to delete Restore1 folder - $($_.Exception.Message)" | Out-File -FilePath $msg_logs -Append
    }
    try {
        $tempFolder = $namespace.Session.Folders.Item("$MyEmail").Folders | Where-Object { $_.Name -eq "Restore2" }
        if ($tempFolder) {
            $tempFolder.Delete()
            "TIME: $(get-dateNow) | INFO | Deleted Restore2 folder after successful import." | Out-File -FilePath $msg_logs -Append
        }
    }
    catch {
        "TIME: $(get-dateNow)  | ERROR |: Failed to delete Restore2 folder - $($_.Exception.Message)" | Out-File -FilePath $msg_logs -Append
    }
}
#DEBUG
