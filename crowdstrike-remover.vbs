'===========================================================================
' CrowdStrike Falcon Corrupted Installation Cleanup Script
' WARNING: Run ONLY when standard uninstallation has failed
' Creates system restore point backup first
'===========================================================================

Option Explicit

Dim objShell, objFSO, logFile, timestamp
Dim intResponse, strComputer, regKey, folderPath

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

timestamp = Replace(Replace(Replace(Now, "/", "-"), ":" , "_"), " ", "")
logFile = objShell.ExpandEnvironmentStrings("%TEMP%") & "\CrowdStrikeCleanup_" & timestamp & ".log"

' Logging function
Sub WriteLog(message)
    Dim fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(logFile, 8, True)
    file.WriteLine Now & " - " & message
    file.Close
End Sub

WriteLog "=========================================="
WriteLog "CrowdStrike Corruption Removal Script Started"
WriteLog "Computer: " & objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
WriteLog "=========================================="

' Confirm user understands risks
intResponse = MsgBox( _
    "CROWDSTRIKE MANUAL REMOVAL SCRIPT" & vbCrLf & vbCrLf & _
    "WARNING: This script will:" & vbCrLf & _
    "- Delete CrowdStrike registry keys" & vbCrLf & _
    "- Remove CrowdStrike program files" & vbCrLf & _
    "- May require system reboot" & vbCrLf & vbCrLf & _
    "REQUIREMENTS:" & vbCrLf & _
    "- Run ONLY after standard uninstall failed" & vbCrLf & _
    "- Have system restore point created first" & vbCrLf & _
    "- Administrator privileges required" & vbCrLf & vbCrLf & _
    "Proceed?", vbYesNo + vbCritical + vbDefaultButton2, "Confirmation Required")

If intResponse <> vbYes Then
    WriteLog "User cancelled operation"
    MsgBox "Operation cancelled. No changes were made.", vbInformation
    WScript.Quit 0
End If

WriteLog "User confirmed - beginning cleanup process"

' Check for administrator privileges
On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c net session >nul 2>&1", 0, True
If Err.Number <> 0 Then
    WriteLog "ERROR: Script must be run as Administrator"
    MsgBox "This script requires Administrator privileges. Please right-click and 'Run as administrator'.", vbCritical
    WScript.Quit 1
End If
On Error GoTo 0

' Registry keys to remove
Dim regKeys(4)
regKeys(0) = "HKLM\SOFTWARE\CrowdStrike"
regKeys(1) = "HKLM\SYSTEM\CrowdStrike"  
regKeys(2) = "HKLM\SYSTEM\CurrentControlSet\Services\CSAgent"
regKeys(3) = "HKLM\SYSTEM\CurrentControlSet\Services\CSAgent\Sim"
regKeys(4) = "HKLM\SYSTEM\CurrentControlSet\Services\CSFalconService"

WriteLog "Beginning registry key removal..."

For Each regKey In regKeys
    On Error Resume Next
    objShell.RegDelete regKey
    If Err.Number = 0 Then
        WriteLog "SUCCESS: Removed " & regKey
    Else
        WriteLog "FAILED: Could not remove " & regKey & " (may not exist or permission issue)"
        Err.Clear
    End If
    On Error GoTo 0
Next

' Remove file directories
Dim fsPath(1)
fsPath(0) = objShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\CrowdStrike"
fsPath(1) = objShell.ExpandEnvironmentStrings("%ProgramData%") & "\CrowdStrike"

WriteLog "Beginning file directory removal..."

' FIXED: Changed loop variable to 'folderPath' to prevent collision with array 'fsPath'
For Each folderPath In fsPath
    On Error Resume Next
    If objFSO.FolderExists(folderPath) Then
        objFSO.DeleteFolder folderPath, True
        If Err.Number = 0 Then
            WriteLog "SUCCESS: Removed folder " & folderPath
        Else
            WriteLog "FAILED: Could not remove folder " & folderPath & " (may be in use)"
            Err.Clear
        End If
    Else
        WriteLog "SKIPPED: Folder does not exist - " & folderPath
    End If
    On Error GoTo 0
Next

WriteLog "Registry and file cleanup completed"
WriteLog "Recommended: Reboot system to ensure all services are cleared"

' Offer reboot option
intResponse = MsgBox("Cleanup process complete. Would you like to reboot now?", _
    vbYesNo + vbQuestion, "Reboot Recommended")

If intResponse = vbYes Then
    WriteLog "Initiating system reboot..."
    objShell.Run "shutdown /r /t 60", 0, False
    MsgBox "System will reboot in 60 seconds. Save all work.", vbInformation
End If

WriteLog "Script completed - Log saved to: " & logFile

MsgBox "Cleanup completed!" & vbCrLf & vbCrLf & _
       "Log file: " & logFile & vbCrLf & _
       "Please verify CrowdStrike components are removed after reboot.", vbInformation

WriteLog "=========================================="
WriteLog "Script execution finished"
WriteLog "=========================================="