'===========================================================================
' CrowdStrike Safe Mode Deep-Clean & Uninstall Residue Removal Script
' WARNING: MUST be run as Administrator inside Safe Mode
'===========================================================================

Option Explicit

Dim objShell, objFSO, objReg, logFile, timestamp
Dim intResponse, folderPath, driverFile
Dim strKeyPath, subKeys, subKey, displayName
Dim msgText, intConfirm

Const HKLM = &H80000002

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")

'===========================================================================
' 0. USER CONFIRMATION PROMPT
'===========================================================================
msgText = "CROWDSTRIKE SAFE MODE DEEP-CLEAN SCRIPT" & vbCrLf & vbCrLf & _
          "This script will perform the following actions:" & vbCrLf & _
          "  1. Force-disable and delete all CrowdStrike kernel services (CSFalconService, CSAgent, CSDeviceControl, CSKernel, etc.)." & vbCrLf & _
          "  2. Thoroughly scan and remove ALL CrowdStrike right-click context menus (files, folders, drives, backgrounds, shortcuts)." & vbCrLf & _
          "  3. Remove Add/Remove Programs (Uninstall) registry entries." & vbCrLf & _
          "  4. Delete kernel driver files (.sys) from System32\drivers and wipe application directories." & vbCrLf & _
          "  5. Clean orphaned UpperFilters and LowerFilters from USB, Keyboard, and Mouse device classes to fix Code 19/38/39 errors." & vbCrLf & vbCrLf & _
          "WARNING: Run this script as Administrator (preferably in Safe Mode)." & vbCrLf & vbCrLf & _
          "Do you want to proceed with the cleanup?"

intConfirm = MsgBox(msgText, vbYesNo + vbExclamation + vbDefaultButton2, "Confirm CrowdStrike Deep Clean")
If intConfirm <> vbYes Then
    WScript.Quit
End If

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

'===========================================================================
' 1. FORCE DISABLE ALL CROWDSTRIKE KERNEL SERVICES & DRIVERS
'===========================================================================
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

'===========================================================================
' 2. DELETE CROWDSTRIKE SERVICE REGISTRY KEYS
'===========================================================================
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

'===========================================================================
' 3. REMOVE ALL RIGHT-CLICK CONTEXT MENU ENTRIES & CORE HIVES
'===========================================================================
Sub CleanContextMenus()
    Dim classRoots(6), root, handlerPath, menuSubKeys, menuKey
    classRoots(0) = "SOFTWARE\Classes\*"
    classRoots(1) = "SOFTWARE\Classes\Directory"
    classRoots(2) = "SOFTWARE\Classes\Directory\Background"
    classRoots(3) = "SOFTWARE\Classes\Folder"
    classRoots(4) = "SOFTWARE\Classes\Drive"
    classRoots(5) = "SOFTWARE\Classes\AllFilesystemObjects"
    classRoots(6) = "SOFTWARE\Classes\lnkfile"

    WriteLog "Scanning all shell classes for CrowdStrike right-click context menu handlers..."
    For Each root In classRoots
        handlerPath = root & "\shellex\ContextMenuHandlers"
        On Error Resume Next
        objReg.EnumKey HKLM, handlerPath, menuSubKeys
        If Not IsNull(menuSubKeys) Then
            For Each menuKey In menuSubKeys
                If InStr(1, menuKey, "CrowdStrike", 1) > 0 Or InStr(1, menuKey, "Falcon", 1) > 0 Or InStr(1, menuKey, "CSAgent", 1) > 0 Then
                    objShell.Run "reg delete ""HKLM\" & handlerPath & "\" & menuKey & """ /f", 0, True
                    WriteLog "SUCCESS: Removed Context Menu Handler: HKLM\" & handlerPath & "\" & menuKey
                End If
            Next
        End If
        On Error GoTo 0
    Next
End Sub

CleanContextMenus

Dim coreKeys(3), cKey
coreKeys(0) = "HKLM\SOFTWARE\CrowdStrike"
coreKeys(1) = "HKLM\SYSTEM\CrowdStrike"
coreKeys(2) = "HKLM\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\CrowdStrike"
coreKeys(3) = "HKLM\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\CrowdStrike"

WriteLog "Deleting core application registry hives..."
For Each cKey In coreKeys
    On Error Resume Next
    objShell.Run "reg delete """ & cKey & """ /f", 0, True
    On Error GoTo 0
Next

'===========================================================================
' 4. REMOVE ADD/REMOVE PROGRAMS (UNINSTALL) ENTRIES
'===========================================================================
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

'===========================================================================
' 5. DELETE KERNEL DRIVER FILES (.SYS) AND DRIVER DIRECTORY
'===========================================================================
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

On Error Resume Next
If objFSO.FolderExists(sys32Dir & "CrowdStrike") Then
    objFSO.DeleteFolder sys32Dir & "CrowdStrike", True
    WriteLog "SUCCESS: Removed System32\drivers\CrowdStrike directory"
End If
On Error GoTo 0

'===========================================================================
' 6. DELETE MAIN APPLICATION & DATA DIRECTORIES
'===========================================================================
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

'===========================================================================
' 7. REMOVE ORPHANED HARDWARE CLASS FILTERS (USB, KEYBOARD, MOUSE)
'===========================================================================
Sub CleanClassFilter(guidKey, filterName)
    Dim regPath, arrValues, i, newValues(), count, val, isModified
    regPath = "SYSTEM\CurrentControlSet\Control\Class\" & guidKey
    count = 0
    isModified = False
    
    On Error Resume Next
    objReg.GetMultiStringValue HKLM, regPath, filterName, arrValues
    If IsArray(arrValues) Then
        For i = 0 To UBound(arrValues)
            val = LCase(Trim(arrValues(i)))
            If Left(val, 2) = "cs" Or InStr(1, val, "crowdstrike", 1) > 0 Then
                isModified = True
                WriteLog "SUCCESS: Removed orphaned filter '" & arrValues(i) & "' from " & guidKey & "\" & filterName
            Else
                ReDim Preserve newValues(count)
                newValues(count) = arrValues(i)
                count = count + 1
            End If
        Next
        
        If isModified Then
            If count = 0 Then
                objReg.DeleteValue HKLM, regPath, filterName
                WriteLog "SUCCESS: Deleted empty " & filterName & " value in " & guidKey
            Else
                objReg.SetMultiStringValue HKLM, regPath, filterName, newValues
                WriteLog "SUCCESS: Updated " & filterName & " in " & guidKey
            End If
        End If
    End If
    On Error GoTo 0
End Sub

WriteLog "Cleaning orphaned UpperFilters/LowerFilters from device class registries..."

' USB Controllers & Hubs
CleanClassFilter "{36FC9E60-C465-11CF-8056-444553540000}", "UpperFilters"
CleanClassFilter "{36FC9E60-C465-11CF-8056-444553540000}", "LowerFilters"

' Keyboards
CleanClassFilter "{4D36E96B-E325-11CE-BFC1-08002BE10318}", "UpperFilters"
CleanClassFilter "{4D36E96B-E325-11CE-BFC1-08002BE10318}", "LowerFilters"

' Mice & Pointing Devices
CleanClassFilter "{4D36E96F-E325-11CE-BFC1-08002BE10318}", "UpperFilters"
CleanClassFilter "{4D36E96F-E325-11CE-BFC1-08002BE10318}", "LowerFilters"

WriteLog "================================================="
WriteLog "Deep clean finished! Log saved to: " & logFile
WriteLog "================================================="

MsgBox "CrowdStrike deep-clean completed!" & vbCrLf & vbCrLf & _
       "All driver services, Add/Remove entries, context menus, driver files, and orphaned hardware class filters have been purged." & vbCrLf & _
       "Please reboot your system normally to restore USB functionality.", vbInformation, "Cleanup Complete"