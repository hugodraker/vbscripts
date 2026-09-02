Option Explicit

' ==============================================================================
' ENFORCE CSCRIPT.EXE ONLY
' ==============================================================================
If InStr(1, LCase(WScript.FullName), "wscript.exe") > 0 Then
    MsgBox "This script must be run using cscript.exe", vbCritical, "Error"
    WScript.Quit
End If

' ==============================================================================
' CONFIGURATION
' ==============================================================================
Dim SERVER_NAME, CONN_STRING
SERVER_NAME   = "localhost"          ' Primary instance

' Build primary connection string
CONN_STRING = "Provider=SQLOLEDB;Data Source=" & SERVER_NAME & ";Initial Catalog=master;Integrated Security=SSPI;"

' ==============================================================================
' MAIN SCRIPT
' ==============================================================================
Dim conn, fso, outFile, rsDBs, sqlDBs, rsTables, sqlTables
Dim currentDB, safeDBName, schemaName, tableName
Dim databasesProcessed : databasesProcessed = 0

Set fso = CreateObject("Scripting.FileSystemObject")
Set conn = CreateObject("ADODB.Connection")

' -------------------------------
' Attempt primary connection
' -------------------------------
On Error Resume Next
conn.Open CONN_STRING

If Err.Number <> 0 Then
    WScript.Echo "Primary connection to '" & SERVER_NAME & "' failed: " & Err.Description
    Err.Clear

    ' -------------------------------
    ' Fallback to SQLEXPRESS
    ' -------------------------------
    WScript.Echo "Attempting fallback connection to .\SQLEXPRESS..."
    CONN_STRING = "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=master;Integrated Security=SSPI;"
    conn.Open CONN_STRING

    If Err.Number <> 0 Then
        WScript.Echo "Fallback connection to .\SQLEXPRESS failed: " & Err.Description
        WScript.Quit
    Else
        SERVER_NAME = ".\SQLEXPRESS"
        WScript.Echo "Connected using fallback instance: .\SQLEXPRESS"
    End If
End If
On Error GoTo 0

WScript.Echo "Connected to SQL Server instance: " & SERVER_NAME
WScript.Echo "Fetching databases..."

' Query to get all user databases that are currently ONLINE
sqlDBs = "SELECT name FROM sys.databases WHERE database_id > 4 AND state_desc = 'ONLINE' ORDER BY name"
Set rsDBs = conn.Execute(sqlDBs)

If rsDBs.EOF Then
    WScript.Echo "No user databases found."
    WScript.Quit
End If

' ==============================================================================
' Loop through each database
' ==============================================================================
Do While Not rsDBs.EOF
    currentDB = rsDBs("name").Value
    WScript.Echo "========================================="
    WScript.Echo "Database: " & currentDB
    
    ' Switch the connection context to the current database
    On Error Resume Next
    conn.Execute "USE [" & currentDB & "]"
    If Err.Number <> 0 Then
        WScript.Echo "  -> Skipping (Cannot access: " & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
    Else
        On Error GoTo 0

        ' Sanitize DB name for valid Windows filename
        safeDBName = SanitizeFileName(currentDB)
        
        ' Create a unique SQL file for this specific database
        Set outFile = fso.CreateTextFile(safeDBName & "_Export.sql", True)
        
        ' Write Database Header to the SQL file
        outFile.WriteLine "-- ========================================================="
        outFile.WriteLine "-- DATABASE: [" & currentDB & "]"
        outFile.WriteLine "-- ========================================================="
        outFile.WriteLine "USE [" & currentDB & "];"
        outFile.WriteLine "GO"
        outFile.WriteLine ""
        
        ' Get all user tables in the CURRENT database
        sqlTables = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_SCHEMA, TABLE_NAME"
        Set rsTables = conn.Execute(sqlTables)
        
        If rsTables.EOF Then
            WScript.Echo "  -> No tables found."
            outFile.WriteLine "-- (No tables found in this database)" & vbCrLf
        Else
            ' Loop through all tables in the current database
            Do While Not rsTables.EOF
                schemaName = rsTables("TABLE_SCHEMA").Value
                tableName = rsTables("TABLE_NAME").Value
                
                WScript.Echo "  -> Processing Table: [" & schemaName & "].[" & tableName & "]"
                
                outFile.WriteLine "-- Table: [" & schemaName & "].[" & tableName & "]"
                
                ' 1. Generate Basic CREATE TABLE
                outFile.WriteLine GenerateCreateTable(conn, schemaName, tableName)
                
                ' 2. Generate ALL Rows INSERT (Writes directly to file)
                GenerateInsertRows conn, schemaName, tableName, outFile
                outFile.WriteLine ""
                
                rsTables.MoveNext
            Loop
        End If
        Set rsTables = Nothing
        
        ' Close the file specific to this database
        outFile.Close
        databasesProcessed = databasesProcessed + 1
    End If
    
    rsDBs.MoveNext
Loop

' ==============================================================================
' Finalization
' ==============================================================================
WScript.Echo "========================================="
If databasesProcessed > 0 Then
    WScript.Echo "Export complete! Created " & databasesProcessed & " individual .sql file(s)."
Else
    WScript.Echo "No databases were successfully processed."
End If

conn.Close
Set rsDBs = Nothing
Set conn = Nothing
Set fso = Nothing

' ==============================================================================
' FUNCTIONS
' ==============================================================================

Function SanitizeFileName(ByVal fName)
    Dim invalidChars, i
    ' Replace characters that are invalid in Windows file names
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = 0 To UBound(invalidChars)
        fName = Replace(fName, invalidChars(i), "_")
    Next
    SanitizeFileName = fName
End Function

Function GenerateCreateTable(conn, schemaName, tableName)
    Dim rsCols, sql, createStr, isFirst
    createStr = "CREATE TABLE [" & schemaName & "].[" & tableName & "] (" & vbCrLf
    
    sql = "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE " & _
          "FROM INFORMATION_SCHEMA.COLUMNS " & _
          "WHERE TABLE_SCHEMA = '" & schemaName & "' AND TABLE_NAME = '" & tableName & "' " & _
          "ORDER BY ORDINAL_POSITION"
          
    Set rsCols = conn.Execute(sql)
    isFirst = True
    
    Do While Not rsCols.EOF
        If Not isFirst Then createStr = createStr & "," & vbCrLf
        isFirst = False
        
        Dim cName, cType, cLen, cPrec, cScale, cNull
        cName = rsCols("COLUMN_NAME").Value
        cType = rsCols("DATA_TYPE").Value
        cLen = rsCols("CHARACTER_MAXIMUM_LENGTH").Value
        cPrec = rsCols("NUMERIC_PRECISION").Value
        cScale = rsCols("NUMERIC_SCALE").Value
        cNull = rsCols("IS_NULLABLE").Value
        
        createStr = createStr & "    [" & cName & "] " & cType
        
        If cType = "varchar" Or cType = "nvarchar" Or cType = "char" Or cType = "nchar" Or cType = "varbinary" Then
            If Not IsNull(cLen) Then
                If cLen = -1 Then
                    createStr = createStr & "(MAX)"
                Else
                    createStr = createStr & "(" & cLen & ")"
                End If
            End If
        ElseIf cType = "decimal" Or cType = "numeric" Then
            If Not IsNull(cPrec) And Not IsNull(cScale) Then
                createStr = createStr & "(" & cPrec & ", " & cScale & ")"
            End If
        End If
        
        If cNull = "NO" Then
            createStr = createStr & " NOT NULL"
        Else
            createStr = createStr & " NULL"
        End If
        
        rsCols.MoveNext
    Loop
    
    createStr = createStr & vbCrLf & ");" & vbCrLf & "GO"
    GenerateCreateTable = createStr
End Function

Sub GenerateInsertRows(conn, schemaName, tableName, outFile)
    Dim rsRows, sql, insertStr, colStr, valStr, fld, vType, vVal
    
    ' Select ALL rows instead of TOP 1
    sql = "SELECT * FROM [" & schemaName & "].[" & tableName & "]"
    
    On Error Resume Next
    Set rsRows = conn.Execute(sql)
    If Err.Number <> 0 Then
        outFile.WriteLine "-- Could not fetch rows (Table might be locked or unselectable)."
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If rsRows.EOF Then
        outFile.WriteLine "-- (Table is empty. No INSERT generated.)"
        Exit Sub
    End If
    
    ' Loop through every record in the table
    Do While Not rsRows.EOF
        colStr = ""
        valStr = ""
        Dim i: i = 0
        
        For Each fld In rsRows.Fields
            If i > 0 Then
                colStr = colStr & ", "
                valStr = valStr & ", "
            End If
            colStr = colStr & "[" & fld.Name & "]"
            
            If IsNull(fld.Value) Then
                valStr = valStr & "NULL"
            Else
                vType = VarType(fld.Value)
                Select Case vType
                    Case 8 ' vbString
                        valStr = valStr & "N'" & Replace(CStr(fld.Value), "'", "''") & "'"
                    Case 7 ' vbDate
                        vVal = fld.Value
                        valStr = valStr & "'" & Year(vVal) & "-" & Right("0" & Month(vVal), 2) & "-" & Right("0" & Day(vVal), 2) & _
                                 " " & Right("0" & Hour(vVal), 2) & ":" & Right("0" & Minute(vVal), 2) & ":" & Right("0" & Second(vVal), 2) & "'"
                    Case 11 ' vbBoolean
                        If fld.Value Then valStr = valStr & "1" Else valStr = valStr & "0"
                    Case 8209 ' vbArray + vbByte (Binary/Image data)
                        valStr = valStr & "NULL /* Binary Data Skipped */"
                    Case Else
                        valStr = valStr & CStr(fld.Value)
                End Select
            End If
            i = i + 1
        Next
        
        insertStr = "INSERT INTO [" & schemaName & "].[" & tableName & "] (" & colStr & ") VALUES (" & valStr & ");"
        outFile.WriteLine insertStr
        
        rsRows.MoveNext
    Loop
    
    outFile.WriteLine "GO"
End Sub