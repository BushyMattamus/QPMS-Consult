Attribute VB_Name = "Access_Connection_Manager"
'=============================================================================
' Module: Access_Connection_Manager
' Purpose: Manage ADODB connections to the Access database
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: GetDatabaseConnection
' Returns: A new ADODB.Connection object to the Access database
'------------------------------------------------------------------------------
Public Function GetDatabaseConnection() As ADODB.Connection
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim dbPath As String
    
    ' Get database path from central resolver
    dbPath = PathResolver.GetDatabasePath()
    
    ' Check if database exists
    If Dir(dbPath) = "" Then
        PathResolver.ShowMessage "Database does not exist at: " & dbPath, vbCritical, "Connection Failed"
        Set GetDatabaseConnection = Nothing
        Exit Function
    End If
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    On Error GoTo ErrorHandler
    
    If Err.Number <> 0 Then
        PathResolver.ShowMessage "Failed to connect to database: " & Err.Description, vbCritical, "Connection Failed"
        Set GetDatabaseConnection = Nothing
        Exit Function
    End If
    
    ' Verify connection is open
    If conn.State <> adStateOpen Then
        PathResolver.ShowMessage "Connection state is not open", vbCritical, "Connection Failed"
        Set GetDatabaseConnection = Nothing
        Exit Function
    End If
    
    Set GetDatabaseConnection = conn
    Exit Function
    
ErrorHandler:
    Set GetDatabaseConnection = Nothing
End Function

'------------------------------------------------------------------------------
' Function: CloseDatabaseConnection
' Purpose: Close an existing database connection
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function CloseDatabaseConnection(conn As ADODB.Connection) As Boolean
    On Error GoTo ErrorHandler
    
    If conn Is Nothing Then
        CloseDatabaseConnection = True
        Exit Function
    End If
    
    If conn.State = adStateOpen Then
        On Error Resume Next
        conn.Close
        If Err.Number <> 0 Then
            PathResolver.ShowMessage "Failed to close database connection: " & Err.Description, vbCritical, "Close Failed"
            CloseDatabaseConnection = False
            Exit Function
        End If
    End If
    
    Set conn = Nothing
    CloseDatabaseConnection = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error closing database connection: " & Err.Description, vbCritical, "Close Failed"
    CloseDatabaseConnection = False
End Function

'------------------------------------------------------------------------------
' Function: ExecuteSQL
' Purpose: Execute a SQL command against the database
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function ExecuteSQL(conn As ADODB.Connection, sql As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    
    On Error Resume Next
    cmd.Execute
    If Err.Number <> 0 Then
        PathResolver.ShowMessage "SQL execution failed: " & Err.Description, vbCritical, "SQL Execution Failed"
        ExecuteSQL = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ExecuteSQL = True
    Exit Function
    
ErrorHandler:
    ExecuteSQL = False
End Function

'------------------------------------------------------------------------------
' Function: ExecuteQueryAndReturnRecordset
' Purpose: Execute a SELECT query and return the recordset
' Returns: ADODB.Recordset object, or Nothing on error
'------------------------------------------------------------------------------
Public Function ExecuteQueryAndReturnRecordset(conn As ADODB.Connection, sql As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Set ExecuteQueryAndReturnRecordset = rs
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Query execution failed: " & Err.Description, vbCritical, "Query Failed"
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set ExecuteQueryAndReturnRecordset = Nothing
End Function

'------------------------------------------------------------------------------
' Function: TestConnection
' Purpose: Test if the database connection is working
' Returns: True if connection is working, False otherwise
'------------------------------------------------------------------------------
Public Function TestConnection() As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    Set conn = GetDatabaseConnection()
    If conn Is Nothing Then
        TestConnection = False
        Exit Function
    End If
    
    sql = "SELECT COUNT(*) AS [RecordCount] FROM [MasterDataFinishedGoods]"
    Set rs = ExecuteQueryAndReturnRecordset(conn, sql)
    
    If rs Is Nothing Then
        TestConnection = False
        Exit Function
    End If
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    TestConnection = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Connection test failed: " & Err.Description, vbCritical, "Connection Test Failed"
    TestConnection = False
End Function

'------------------------------------------------------------------------------
' Function: GetTableNames
' Purpose: Get a list of all non-system table names in the database
' Returns: Collection of table names
'------------------------------------------------------------------------------
Public Function GetTableNames(conn As ADODB.Connection) As Collection
    On Error GoTo ErrorHandler
    
    Dim rsTables As ADODB.Recordset
    Dim tableName As String
    Dim tables As Collection
    Dim i As Integer
    
    Set rsTables = New ADODB.Recordset
    rsTables.Open "SELECT [Name] FROM [MSysObjects] WHERE [Type]=-32768 AND [Name] NOT LIKE 'MSys*'", _
                  conn, adOpenForwardOnly, adLockOptimistic
    
    Set tables = New Collection
    i = 1
    Do While Not rsTables.EOF
        tableName = rsTables!Name
        If Left(tableName, 4) <> "MSys" Then
            tables.Add tableName, CStr(i)
        End If
        i = i + 1
        rsTables.MoveNext
    Loop
    
    rsTables.Close
    Set rsTables = Nothing
    
    Set GetTableNames = tables
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error getting table names: " & Err.Description, vbCritical, "Get Table Names Failed"
    Set GetTableNames = Nothing
End Function

'------------------------------------------------------------------------------
' Function: GetRecordCountForTable
' Purpose: Get the record count for a specific table
' Returns: Number of records, or -1 on error
'------------------------------------------------------------------------------
Public Function GetRecordCountForTable(conn As ADODB.Connection, tableName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "SELECT COUNT(*) AS [Count] FROM [" & tableName & "]"
    Set rs = ExecuteQueryAndReturnRecordset(conn, sql)
    
    If rs Is Nothing Then
        GetRecordCountForTable = -1
        Exit Function
    End If
    
    GetRecordCountForTable = rs!Count
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrorHandler:
    GetRecordCountForTable = -1
End Function
