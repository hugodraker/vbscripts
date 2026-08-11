Option Explicit

' ==============================================================================
' CONFIGURATION
' ==============================================================================
Dim SERVER_NAME, OUTPUT_FILE, CONN_STRING
SERVER_NAME   = "localhost"          ' Change if using a named instance
OUTPUT_FILE   = "Exported_All_Databases_Tables.sql"

' Connect to 'master' so we can query sys.databases first
CONN_STRING = "Provider=SQLOLEDB;Data Source=" & SERVER_NAME & ";Initial Catalog=master;Integrated Security=SSPI;"

' ==============================================================================
' MAIN SCRIPT
' ==============================================================================
Dim conn, fso, outFile, rsDBs, sqlDBs, rsTables, sqlTables
Dim currentDB, schemaName, tableName

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile(OUTPUT_FILE, True)

Set conn = CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open CONN_STRING
If Err.Number <> 0 Then
    WScript.Echo "Error connecting to SQL Server: " & Err.Description
    WScript.Quit
End If
On Error GoTo 0

WScript.Echo "Connected to " & SERVER_NAME & ". Fetching databases..."

' Query to get all user databases that are currently ONLINE
sqlDBs = "SELECT name FROM sys.databases WHERE database_id > 4 AND state_desc = 'ONLINE' ORDER BY name"
Set rsDBs = conn.Execute(sqlDBs)

If rsDBs.EOF Then
    WScript.Echo "No user databases found."
    WScript.Quit
End If

' Loop through each database
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
                
                ' 2. Generate 1 Row INSERT
                outFile.WriteLine GenerateInsertRow(conn, schemaName, tableName)
                outFile.WriteLine ""
                
                rsTables.MoveNext
            Loop
        End If
        Set rsTables = Nothing
    End If
    
    rsDBs.MoveNext
Loop

outFile.Close
conn.Close
Set rsDBs = Nothing
Set conn = Nothing
Set fso = Nothing

WScript.Echo "========================================="
WScript.Echo "Export complete! File saved to: " & OUTPUT_FILE

' ==============================================================================
' FUNCTIONS
' ==============================================================================

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
        
        ' Handle Lengths and Precision
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
        
        ' Handle Nullability
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

Function GenerateInsertRow(conn, schemaName, tableName)
    Dim rsRow, sql, insertStr, colStr, valStr, fld, vType, vVal
    sql = "SELECT TOP 1 * FROM [" & schemaName & "].[" & tableName & "]"
    
    On Error Resume Next
    Set rsRow = conn.Execute(sql)
    If Err.Number <> 0 Then
        GenerateInsertRow = "-- Could not fetch row (Table might be locked or unselectable)."
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    If rsRow.EOF Then
        GenerateInsertRow = "-- (Table is empty. No INSERT generated.)"
        Exit Function
    End If
    
    colStr = ""
    valStr = ""
    Dim i: i = 0
    
    For Each fld In rsRow.Fields
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
                    ' Format Date as safe ISO format: YYYY-MM-DD HH:MM:SS
                    vVal = fld.Value
                    valStr = valStr & "'" & Year(vVal) & "-" & Right("0" & Month(vVal), 2) & "-" & Right("0" & Day(vVal), 2) & " " & Right("0" & Hour(vVal), 2) & ":" & Right("0" & Minute(vVal), 2) & ":" & Right("0" & Second(vVal), 2) & "'"
                Case 11 ' vbBoolean
                    If fld.Value Then valStr = valStr & "1" Else valStr = valStr & "0"
                Case 8209 ' vbArray + vbByte (Binary/Image data)
                    valStr = valStr & "NULL /* Binary Data Skipped */"
                Case Else
                    ' Numeric and other fallback
                    valStr = valStr & CStr(fld.Value)
            End Select
        End If
        i = i + 1
    Next
    
    insertStr = "INSERT INTO [" & schemaName & "].[" & tableName & "] (" & colStr & ") VALUES (" & valStr & ");" & vbCrLf & "GO"
    GenerateInsertRow = insertStr
End Function