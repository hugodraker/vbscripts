Option Explicit

' ==============================================================================
' ENFORCE CSCRIPT.EXE ONLY
' ==============================================================================
If InStr(1, LCase(WScript.FullName), "wscript.exe") > 0 Then
    MsgBox "This script must be run using cscript.exe", vbCritical, "Error"
    WScript.Quit
End If

WScript.Echo "SQL Maintenance Script Starting..."
WScript.Echo ""

' ==============================================================================
' CONFIGURATION
' ==============================================================================
Dim SERVER_NAME, CONN_STRING
SERVER_NAME = "localhost"

CONN_STRING = "Provider=SQLOLEDB;Data Source=" & SERVER_NAME & ";Initial Catalog=master;Integrated Security=SSPI;"

Dim conn
Set conn = CreateObject("ADODB.Connection")

' ==============================================================================
' CONNECT WITH FALLBACK
' ==============================================================================
On Error Resume Next
conn.Open CONN_STRING

If Err.Number <> 0 Then
    WScript.Echo "Primary connection failed: " & Err.Description
    Err.Clear

    WScript.Echo "Trying fallback instance .\SQLEXPRESS..."
    CONN_STRING = "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=master;Integrated Security=SSPI;"
    conn.Open CONN_STRING

    If Err.Number <> 0 Then
        WScript.Echo "Fallback connection failed: " & Err.Description
        WScript.Quit
    Else
        SERVER_NAME = ".\SQLEXPRESS"
        WScript.Echo "Connected to fallback instance."
    End If
End If
On Error GoTo 0

WScript.Echo "Connected to SQL Server: " & SERVER_NAME
WScript.Echo ""

' ==============================================================================
' GET USER DATABASES
' ==============================================================================
Dim rsDBs, sqlDBs
sqlDBs = "SELECT name FROM sys.databases WHERE database_id > 4 AND state_desc='ONLINE' ORDER BY name"
Set rsDBs = conn.Execute(sqlDBs)

If rsDBs.EOF Then
    WScript.Echo "No user databases found."
    WScript.Quit
End If

' ==============================================================================
' MAINTENANCE LOOP
' ==============================================================================
Do While Not rsDBs.EOF
    Dim dbName
    dbName = rsDBs("name").Value

    WScript.Echo "==============================================="
    WScript.Echo "Processing database: " & dbName

    ' Switch DB
    conn.Execute "USE [" & dbName & "]"

    ' ------------------------------
    ' CLEANUP
    ' ------------------------------
    WScript.Echo " - Cleanup: Updating statistics..."
    conn.Execute "EXEC sp_updatestats"

    WScript.Echo " - Cleanup: Clearing cache..."
    conn.Execute "DBCC FREEPROCCACHE"
    conn.Execute "DBCC DROPCLEANBUFFERS"

    ' ------------------------------
    ' REINDEX
    ' ------------------------------
    WScript.Echo " - Reindexing tables..."

    Dim sqlIndex
    sqlIndex = "DECLARE @tbl NVARCHAR(255);" & _
               "DECLARE c CURSOR FOR " & _
               "SELECT QUOTENAME(s.name) + '.' + QUOTENAME(t.name) " & _
               "FROM sys.tables t JOIN sys.schemas s ON t.schema_id=s.schema_id;" & _
               "OPEN c; FETCH NEXT FROM c INTO @tbl;" & _
               "WHILE @@FETCH_STATUS=0 BEGIN " & _
               "PRINT 'Reindexing ' + @tbl;" & _
               "EXEC('ALTER INDEX ALL ON ' + @tbl + ' REBUILD');" & _
               "FETCH NEXT FROM c INTO @tbl;" & _
               "END; CLOSE c; DEALLOCATE c;"

    conn.Execute sqlIndex

    ' ------------------------------
    ' COMPACT (SAFE SHRINK)
    ' ------------------------------
    WScript.Echo " - Compacting database..."

    Dim sqlShrink
    sqlShrink = "DBCC SHRINKDATABASE ([" & dbName & "], 10)"

    conn.Execute sqlShrink

    WScript.Echo "Finished: " & dbName
    WScript.Echo ""

    rsDBs.MoveNext
Loop

WScript.Echo "==============================================="
WScript.Echo "Maintenance complete on all databases."
WScript.Echo ""

conn.Close
Set conn = Nothing
Set rsDBs = Nothing
