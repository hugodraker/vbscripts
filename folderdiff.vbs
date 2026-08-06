Option Explicit

' Check if running under CScript (Console Mode).
' If not (e.g., double-clicked via WScript), open a new command prompt window,
' forward any command-line arguments, and run with CScript.
If UCase(Right(WScript.FullName, 11)) <> "CSCRIPT.EXE" Then
    Dim objShell, argString, i
    argString = ""
    For i = 0 To WScript.Arguments.Count - 1
        argString = argString & " """ & WScript.Arguments(i) & """"
    Next

    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd.exe /k cscript.exe //nologo """ & WScript.ScriptFullName & """" & argString, 1, False
    WScript.Quit(0)
End If

Dim fso, sourceDir, targetDir

' --- Default Hardcoded Folders ---
sourceDir = "D:\"
targetDir = "C:\QDR\DATA"

' --- Override with Command Line Arguments if Specified ---
If WScript.Arguments.Count >= 1 Then
    sourceDir = WScript.Arguments(0)
End If
If WScript.Arguments.Count >= 2 Then
    targetDir = WScript.Arguments(1)
End If

Set fso = CreateObject("Scripting.FileSystemObject")

' Start recursive check from the source directory
If fso.FolderExists(sourceDir) Then
    FindMissingFiles fso.GetFolder(sourceDir)
Else
    WScript.Echo "Error: Source folder '" & sourceDir & "' was not found."
End If

' Exit cleanly
WScript.Quit(0)


' --- Subroutines ---

Sub FindMissingFiles(folder)
    On Error Resume Next ' Skip system folders with permission errors (e.g., $RECYCLE.BIN)
    
    Dim file, subFolder, targetFilePath
    
    ' Process all files in the current folder
    For Each file In folder.Files
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        
        ' Safely combine target folder path and filename (handles trailing slashes automatically)
        targetFilePath = fso.BuildPath(targetDir, file.Name)
        
        ' Check if the file from the source folder does NOT exist in the target folder
        If Not fso.FileExists(targetFilePath) Then
            ' Print filename and modified date on the same line, separated by a tab
            WScript.Echo file.Name & vbTab & file.DateLastModified
        End If
    Next
    
    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        If Err.Number = 0 Then
            FindMissingFiles subFolder
        End If
        Err.Clear
    Next
    
    On Error GoTo 0
End Sub