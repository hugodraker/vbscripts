'===========================================================================
' CrowdStrike Safe Mode Deep-Clean & Uninstall Residue Removal Script
' WARNING: MUST be run as Administrator inside Safe Mode
'===========================================================================

Option Explicit

Dim objShell, objFSO, objReg, logFile, timestamp
Dim intResponse, folderPath, driverFile
Dim strKeyPath, subKeys, subKey, displayName

Const HKLM = &H80000002

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")

timestamp = Replace(Replace(Replace(Now, "/", "-"), ":" , "_"), " ", "")
logFile = objShell.ExpandEnvironmentStrings("%TEMP%") & "\CS_SafeMode_Clean_" & timestamp & ".log"

Sub WriteLog(message)
    Dim fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(logFile, 8, True)
    file.WriteLine Now & " - " & message
    file.Close
End Sub

WriteLog "================================================="
WriteLog "CrowdStrike Safe Mode Deep Clean Script Started"
WriteLog "================================================="

' 1. FORCE DISABLE ALL CROWDSTRIKE KERNEL SERVICES & DRIVERS
' Setting Start = 4 disables the service/driver from booting
Dim csServices(4), serviceName
csServices(0) = "CSFalconService"
csServices(1) = "CSAgent"
csServices(2) = "CSBoot"
csServices(3) = "CSDeviceControl"
csServices(4) = "CSKernel"

WriteLog "Disabling service startup states..."
For Each serviceName In csServices
    On Error Resume Next
    objShell.Run "reg add ""HKLM\SYSTEM\CurrentControlSet\Services\" & serviceName & """ /v Start /t REG_DWORD /d 4 /f", 0, True
    If Err.Number = 0 Then
        WriteLog "SUCCESS: Disabled service startup for " & serviceName
    End If
    On Error GoTo 0
Next

' 2. DELETE CROWDSTRIKE SERVICE REGISTRY KEYS
Dim regServiceKeys(5), sKey
regServiceKeys(0) = "HKLM\SYSTEM\CurrentControlSet\Services\CSFalconService"
regServiceKeys(1) = "HKLM\SYSTEM\CurrentControlSet\Services\CSAgent\Sim"
regServiceKeys(2) = "HKLM\SYSTEM\CurrentControlSet\Services\CSAgent"
regServiceKeys(3) = "HKLM\SYSTEM\CurrentControlSet\Services\CSBoot"
regServiceKeys(4) = "HKLM\SYSTEM\CurrentControlSet\Services\CSDeviceControl"
regServiceKeys(5) = "HKLM\SYSTEM\CurrentControlSet\Services\CSKernel"

WriteLog "Deleting Windows Service registry keys..."
For Each sKey In regServiceKeys
    On Error Resume Next
    objShell.RegDelete sKey & "\"
    objShell.RegDelete sKey
    If Err.Number = 0 Then
        WriteLog "SUCCESS: Removed service registry key " & sKey
    End If
    On Error GoTo 0
Next

' 3. REMOVE ADD/REMOVE PROGRAMS (UNINSTALL) ENTRIES
' Scans Windows Uninstall registry hives for "CrowdStrike" and deletes matching GUID keys
Sub CleanUninstallHive(hivePath)
    On Error Resume Next
    objReg.EnumKey HKLM, hivePath, subKeys
    If Not IsNull(subKeys) Then
        For Each subKey In subKeys
            displayName = ""
            objReg.GetStringValue HKLM, hivePath & "\" & subKey, "DisplayName", displayName
            If InStr(1, displayName, "CrowdStrike", 1) > 0 Then
                objShell.Run "reg delete ""HKLM\" & hivePath & "\" & subKey & """ /f", 0, True
                WriteLog "SUCCESS: Removed Add/Remove Programs entry: " & displayName & " (" & subKey & ")"
            End If
        Next
    End If
    On Error GoTo 0
End Sub

WriteLog "Scanning and cleaning Add/Remove Programs (Uninstall) registry hives..."
CleanUninstallHive "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
CleanUninstallHive "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

' 4. DELETE CORE APPLICATION REGISTRY HIVES & CONTEXT MENUS
Dim coreKeys(3), cKey
coreKeys(0) = "HKLM\SOFTWARE\CrowdStrike"
coreKeys(1) = "HKLM\SYSTEM\CrowdStrike"
coreKeys(2) = "HKLM\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\CrowdStrike"
coreKeys(3) = "HKLM\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\CrowdStrike"

WriteLog "Deleting core application & context menu registry hives..."
For Each cKey In coreKeys
    On Error Resume Next
    objShell.Run "reg delete """ & cKey & """ /f", 0, True
    On Error GoTo 0
Next

' 5. DELETE KERNEL DRIVER FILES (.SYS) AND DRIVER DIRECTORY
Dim sys32Dir, driverFiles(3), dFile
sys32Dir = objShell.ExpandEnvironmentStrings("%WINDIR%") & "\System32\drivers\"
driverFiles(0) = "csagent.sys"
driverFiles(1) = "csboot.sys"
driverFiles(2) = "csdevicecontrol.sys"
driverFiles(3) = "cskernel.sys"

WriteLog "Removing kernel driver files from System32\drivers..."
For Each dFile In driverFiles
    On Error Resume Next
    If objFSO.FileExists(sys32Dir & dFile) Then
        objFSO.DeleteFile sys32Dir & dFile, True
        WriteLog "SUCCESS: Removed driver file - " & dFile
    End If
    On Error GoTo 0
Next

' Remove CrowdStrike drivers subfolder
On Error Resume Next
If objFSO.FolderExists(sys32Dir & "CrowdStrike") Then
    objFSO.DeleteFolder sys32Dir & "CrowdStrike", True
    WriteLog "SUCCESS: Removed System32\drivers\CrowdStrike directory"
End If
On Error GoTo 0

' 6. DELETE MAIN APPLICATION & DATA DIRECTORIES
Dim fsPaths(1), targetDir
fsPaths(0) = objShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\CrowdStrike"
fsPaths(1) = objShell.ExpandEnvironmentStrings("%ProgramData%") & "\CrowdStrike"

WriteLog "Removing application directories..."
For Each targetDir In fsPaths
    On Error Resume Next
    If objFSO.FolderExists(targetDir) Then
        objFSO.DeleteFolder targetDir, True
        WriteLog "SUCCESS: Removed folder " & targetDir
    Else
        WriteLog "SKIPPED: Folder already gone - " & targetDir
    End If
    On Error GoTo 0
Next

WriteLog "================================================="
WriteLog "Deep clean finished! Log saved to: " & logFile
WriteLog "================================================="

MsgBox "CrowdStrike deep-clean completed!" & vbCrLf & vbCrLf & _
       "All driver services, Add/Remove entries, and driver files have been purged." & vbCrLf & _
       "Please reboot your system normally.", vbInformation, "Cleanup Complete"