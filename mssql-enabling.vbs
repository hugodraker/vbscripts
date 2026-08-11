' =====================================================================
' SQL SERVER LOCALHOST FIX SCRIPT
' Automatically enables TCP/IP, Named Pipes, and SQL Browser
' =====================================================================

' 1. Request Administrator Privileges
If Not WScript.Arguments.Named.Exists("elevated") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
  WScript.Quit
End If

Dim objShell, wmiCIMv2, wmiSQL, sqlNamespace, namespaceFound
Set objShell = CreateObject("WScript.Shell")
strComputer = "."

WScript.Echo "Starting SQL Server Configuration Fix..."

' 2. Find the correct SQL Server WMI Namespace
' (Checks versions from SQL Server 2022 down to 2008)
Dim versions, ver, strPath
versions = Array("16", "15", "14", "13", "12", "11", "10")
namespaceFound = False

For Each ver In versions
    On Error Resume Next
    strPath = "winmgmts:\\" & strComputer & "\root\Microsoft\SqlServer\ComputerManagement" & ver
    Set wmiSQL = GetObject(strPath)
    If Err.Number = 0 Then
        sqlNamespace = strPath
        namespaceFound = True
        Err.Clear
        Exit For
    End If
    Err.Clear
    On Error GoTo 0
Next

If Not namespaceFound Then
    WScript.Echo "ERROR: Could not find SQL Server WMI Provider." & vbCrLf & "Ensure SQL Server is actually installed on this machine."
    WScript.Quit
End If

' 3. Enable TCP/IP and Named Pipes
WScript.Echo "Found SQL WMI Provider: ComputerManagement" & ver & vbCrLf & "Configuring Network Protocols..."
On Error Resume Next
Dim colProtocols, objProtocol
Set colProtocols = wmiSQL.ExecQuery("SELECT * FROM ServerNetworkProtocol WHERE ProtocolName = 'Tcp' OR ProtocolName = 'Np'")

If Err.Number = 0 Then
    For Each objProtocol In colProtocols
        objProtocol.SetEnable()
    Next
    WScript.Echo "TCP/IP and Named Pipes have been enabled."
Else
    WScript.Echo "Warning: Could not configure protocols automatically. Error: " & Err.Description
End If
On Error GoTo 0

' 4. Configure and Start SQL Server Browser Service
Set wmiCIMv2 = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Dim colServices, objService
Set colServices = wmiCIMv2.ExecQuery("Select * from Win32_Service Where Name = 'SQLBrowser'")

For Each objService in colServices
    ' Set to Automatic Start
    objService.ChangeStartMode("Automatic")
    If objService.State <> "Running" Then
        objService.StartService()
        WScript.Echo "SQL Server Browser service started."
    Else
        WScript.Echo "SQL Server Browser service is already running."
    End If
Next

' 5. Restart all SQL Server Engine Services
WScript.Echo vbCrLf & "Restarting SQL Server services to apply changes. Please wait..."
Dim colSQLServices
Set colSQLServices = wmiCIMv2.ExecQuery("Select * from Win32_Service Where Name Like 'MSSQL$%' Or Name = 'MSSQLSERVER'")

For Each objService in colSQLServices
    WScript.Echo "Restarting instance: " & objService.DisplayName & "..."
    ' Using cmd.exe net stop/start ensures the script waits for the service to fully restart
    objShell.Run "cmd.exe /c net stop """ & objService.Name & """ /y", 0, True
    objShell.Run "cmd.exe /c net start """ & objService.Name & """", 0, True
Next

WScript.Echo vbCrLf & "Configuration complete! Your local SQL Server is now set up to accept VBScript connections."