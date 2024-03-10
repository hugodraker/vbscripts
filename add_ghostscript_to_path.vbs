'LICENCE: PUBLIC DOMAIN
'NO WARRANTY
Dim arg, gswinexe, gspath
Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")

if WScript.Arguments.Named.Exists("elevated")=False AND IsSystemUser=False Then 
    CreateObject ("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """"&" /elevated", "","runas", 1
    WScript.Quit
End If

'On error resume next 'comment out to debug
gspath = ReadInstallPath("ghostscript")

If gspath=False then
    Msgbox "Can't find ghostscript path from installed programs, quitting"
    WScript.Quit
End If

WriteEnvPath gspath & "\bin;" & gspath & "\lib;"
oShell.CurrentDirectory = gspath & "\bin"

Set oShell = Nothing
Set fso = Nothing


Function FileExists(FilePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
        FileExists=CBool(1)
    Else
        FileExists=CBool(0)
    End If
End Function

Function ReadInstallPath(strKeyName)
    Dim oReg, oFSO 
    Dim UninstallString, ProductCode
    Dim strComputer, colItems, objWMIService, objItem
    Dim strKeyPath(1), subkey, arrSubKeys
    strComputer = "." 
    'InputBox(UninstallString)

    Const HKEY_LOCAL_MACHINE = &H80000002
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

    strKeyPath(0) = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    strKeyPath(1) = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

    For i=0 to 1
         if oReg.EnumKey(HKEY_LOCAL_MACHINE, strKeyPath(i), arrSubKeys) = 0 then
         'if err.number <> 0 Then
            For Each subkey In arrSubKeys 
               IF instr(1,subkey,strKeyName,1) Then
               str = ReadfromRegistry("HKLM" & "\" & strKeyPath(i) & "\" & subkey & "\" &"UninstallString","1")
               With CreateObject("scripting.filesystemobject")
                   str = Replace(str,"""","")
                   wFolder = .GetParentFolderName(str)
                   ReadInstallPath=wFolder
                 End With
               End If
            Next
        End If
    Next

    Set oReg = Nothing
End Function

Function ReadFromRegistry(strRegistryKey, strDefault)
    Dim value
    On Error Resume Next
    value = oShell.RegRead( strRegistryKey )
    If err.number <> 0 Then
        readFromRegistry = "" 'strDefault
    Else
        readFromRegistry = value
    End If
End Function

Function WriteEnvPath(myKeyValue)
    Const HKEY_LOCAL_MACHINE = &H80000002
    Dim myKeyType, myKey, readFromRegistry, result
    Set WshShell = CreateObject("WScript.Shell")
    myKeyType = "REG_SZ"
    'On Error Resume Next
    myKey = "HKLM\SYSTEM\ControlSet001\Control\Session Manager\Environment\Path"

    path = WshShell.RegRead(myKey)
    If err.number <> 0 Then
        readFromRegistry = ""
        msgbox "Error writing path in registry, please run as administrator"
    Else
        readFromRegistry = path
    End If

    If InStr(readFromRegistry,myKeyValue) = 0 Then 
       if IsSystemUser = 1 Then 
           result=vbYes 'logout without asking, sorry
       Else
           result = msgbox("This should be runas administrator," &vbCRLF &"Add ghostscript to path and logout to update path?", vbYesNo+vbQuestion, "Logout?")
       End if
       Select Case result
            Case vbYes
            WshShell.RegWrite myKey, Replace(readFromRegistry & ";",";;",";") & myKeyValue,myKeyType
            WshShell.Run "shutdown.exe -l",0,True 'doesn't logout SYSTEM
            'oShell.Run  "taskkill /f /im explorer.exe",0,True
            'oShell.Run  "explorer.exe",0,True 'doesnt update current Path Environment Variable
       End Select
    End if
End Function

Function IsSystemUser
    Dim strUser
    strUser = CreateObject("WScript.Network").UserName
    Select Case strUser
    Case "SYSTEM"
        IsSystemUser = 1
    Case Else
        IsSystemUser = 0
    End Select
    'strUser = Nothing
End Function