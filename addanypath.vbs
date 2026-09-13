Option Explicit

Dim objShell, objShellApp, objEnv
Dim strCurrentPath, objFolder, strCustomPath

Set objShell = CreateObject("WScript.Shell")
Set objShellApp = CreateObject("Shell.Application")

' Using "User" environment variables avoids needing admin rights
Set objEnv = objShell.Environment("User")

' Prompt user to browse for any folder
Set objFolder = objShellApp.BrowseForFolder(0, "Select a folder to add to your user PATH:", &H0001 + &H0010, "")

If Not objFolder Is Nothing Then
    strCustomPath = objFolder.Self.Path
    strCurrentPath = objEnv("PATH")
    
    ' Check if the custom path is already in the path to avoid duplicates
    If InStr(1, strCurrentPath, strCustomPath, vbTextCompare) = 0 Then
        If strCurrentPath <> "" And Right(strCurrentPath, 1) <> ";" Then
            strCurrentPath = strCurrentPath & ";"
        End If
        
        ' Append the selected custom path
        objEnv("PATH") = strCurrentPath & strCustomPath & ";"
        
        WScript.Echo "Success! The following path has been added to your user PATH:" & vbCrLf & strCustomPath & vbCrLf & vbCrLf & _
                     "Please restart any open command prompts or terminals for the changes to take effect."
    Else
        WScript.Echo "The selected path is already present in your user PATH." & vbCrLf & strCustomPath & vbCrLf & vbCrLf & _
                     "No changes were made."
    End If
Else
    WScript.Echo "Operation cancelled. No changes were made."
End If