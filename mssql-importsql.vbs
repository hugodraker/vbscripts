Option Explicit

' ==============================================================================
' ENFORCE CSCRIPT.EXE ONLY
' ==============================================================================
If InStr(1, LCase(WScript.FullName), "wscript.exe") > 0 Then
    MsgBox "This script must be run using cscript.exe", vbCritical, "Error"
    WScript.Quit
End If

' ==============================================================================
' CONFIGURATION & CONNECTION
' ==============================================================================
Dim SERVER_NAME, CONN_STRING
SERVER_NAME   = "localhost"

CONN_STRING = "Provider=SQLOLEDB;Data Source=" & SERVER_NAME & ";Initial Catalog=master;Integrated Security=SSPI;"

Dim conn, fso, scriptDir, folder, file
Dim filesProcessed : filesProcessed = 0

Set fso = CreateObject("Scripting.FileSystemObject")
Set conn = CreateObject("ADODB.Connection")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' -------------------------------
' Attempt primary connection
' -------------------------------
On Error Resume Next
conn.Open CONN_STRING

If Err.Number <> 0 Then
    WScript.Echo "Primary connection to '" & SERVER_NAME & "' failed: " & Err.Description
    Err.Clear

    WScript.Echo "Attempting fallback connection to .\SQLEXPRESS..."
    CONN_STRING = "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=master;Integrated Security=SSPI;"
    conn.Open CONN_STRING

    If Err.Number <> 0 Then
        WScript.Echo "Fallback connection failed: " & Err.Description
        WScript.Quit
    Else
        SERVER_NAME = ".\SQLEXPRESS"
        WScript.Echo "Connected using fallback instance: .\SQLEXPRESS"
    End If
End If
On Error GoTo 0

WScript.Echo "Connected to SQL Server: " & SERVER_NAME
WScript.Echo "Scanning directory: " & scriptDir

' ==============================================================================
' LOOP THROUGH .SQL FILES
' ==============================================================================
Set folder = fso.GetFolder(scriptDir)

For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "sql" Then
        WScript.Echo "========================================="
        WScript.Echo "Importing: " & file.Name
        ProcessSqlFile conn, file.Path, fso
        filesProcessed = filesProcessed + 1
    End If
Next

WScript.Echo "========================================="
If filesProcessed > 0 Then
    WScript.Echo "Import complete! Processed " & filesProcessed & " .sql file(s)."
Else
    WScript.Echo "No .sql files were found in the script directory."
End If

conn.Close
Set conn = Nothing
Set folder = Nothing
Set fso = Nothing

' ==============================================================================
' FUNCTIONS
' ==============================================================================

Sub ProcessSqlFile(dbConn, filePath, fileSys)
    Dim txtFile, sqlText, regex, matches, dbName
    Dim batches, batch, i
    
    ' Read entire SQL file
    On Error Resume Next
    Set txtFile = fileSys.OpenTextFile(filePath, 1) ' 1 = ForReading
    If Err.Number <> 0 Then
        WScript.Echo "  -> Error reading file: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If txtFile.AtEndOfStream Then
        WScript.Echo "  -> File is empty."
        txtFile.Close
        Exit Sub
    End If
    
    sqlText = txtFile.ReadAll()
    txtFile.Close
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.MultiLine = True
    
    ' 1. Auto-Create Database if it doesn't exist
    ' Looks for: USE [DatabaseName];
    regex.Pattern = "^USE\s+\[([^\]]+)\]\s*;"
    Set matches = regex.Execute(sqlText)
    If matches.Count > 0 Then
        dbName = matches(0).SubMatches(0)
        WScript.Echo "  -> Target Database: [" & dbName & "]"
        
        Dim checkSql
        checkSql = "IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'" & Replace(dbName, "'", "''") & "') CREATE DATABASE [" & dbName & "]"
        On Error Resume Next
        dbConn.Execute checkSql
        If Err.Number <> 0 Then
            WScript.Echo "  -> Warning: Could not auto-create database. " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' 2. Split file by GO statements for ADODB execution
    regex.Pattern = "^\s*GO\s*$"
    sqlText = regex.Replace(sqlText, "___BATCH_SEPARATOR___")
    
    batches = Split(sqlText, "___BATCH_SEPARATOR___")
    
    ' 3. Execute each batch
    For i = 0 To UBound(batches)
        batch = Trim(batches(i))
        If Len(batch) > 0 Then
            On Error Resume Next
            dbConn.Execute batch
            If Err.Number <> 0 Then
                WScript.Echo "  -> Error in batch execution: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
    
    WScript.Echo "  -> File imported successfully."
End Sub