Attribute VB_Name = "Database_Initialization"
'=============================================================================
' Module: Database_Initialization
' Purpose: Orchestrates the complete database creation and initialization workflow
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: InitializeDatabase
' Purpose: Complete workflow to create and initialize the Access database
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function InitializeDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    Dim dbPath As String
    Dim importCount As Long
    Dim cntFG As Long, cntAT As Long, cntRM As Long, cntRMC As Long
    
    ' Get database path from central resolver
    dbPath = PathResolver.GetDatabasePath()
    
    ' Step 1: Create the database file
    If Access_Database_Creation.CreateEmptyDatabase(dbPath) Then
        ' Step 2: Create all tables (drops existing tables first for clean build)
        If Access_Table_Definitions.CreateAllTables(dbPath) Then
            ' Step 3: Create views (optional)
            Access_Relationships.CreateAllViews dbPath
            
            ' Step 4: Import CSV data into tables
            cntFG = Access_ImportMasterData.ImportMasterDataFinishedGoods(dbPath)
            cntAT = Access_ImportMasterData.ImportAuditTrail(dbPath)
            cntRM = Access_ImportRawMaterial.ImportRawMaterialMaster(dbPath)
            cntRMC = Access_ImportRawMaterial.ImportRMChanges(dbPath)
            
            If cntFG < 0 Or cntAT < 0 Or cntRM < 0 Or cntRMC < 0 Then
                PathResolver.ShowMessage "Failed to import CSV data.", vbCritical, "Initialization Failed"
                InitializeDatabase = False
                Exit Function
            End If
            
            importCount = cntFG + cntAT + cntRM + cntRMC
            
            ' Step 5: Test connection
            If Access_Connection_Manager.TestConnection() Then
                PathResolver.ShowMessage "Database initialization completed successfully! Records imported: " & importCount, vbInformation, "Initialization Complete"
                InitializeDatabase = True
                Exit Function
            Else
                PathResolver.ShowMessage "Failed to test database connection.", vbCritical, "Initialization Failed"
                InitializeDatabase = False
                Exit Function
            End If
        Else
            PathResolver.ShowMessage "Failed to create tables. Aborting initialization.", vbCritical, "Initialization Failed"
            InitializeDatabase = False
            Exit Function
        End If
    Else
        PathResolver.ShowMessage "Failed to create database file. Aborting initialization.", vbCritical, "Initialization Failed"
        InitializeDatabase = False
        Exit Function
    End If
    
ErrorHandler:
    PathResolver.ShowMessage "Error during database initialization: " & Err.Description, vbCritical, "Initialization Failed"
    InitializeDatabase = False
End Function

'------------------------------------------------------------------------------
' Function: ReinitializeDatabase
' Purpose: Completely recreate the database (delete existing, create new)
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function ReinitializeDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    Dim dbPath As String
    Dim importCount As Long
    Dim cntFG As Long, cntAT As Long, cntRM As Long, cntRMC As Long
    
    ' Get database path
    dbPath = PathResolver.GetDatabasePath()
    
    ' Step 1: Delete existing database if it exists
    If Dir(dbPath) <> "" Then
        Access_Database_Creation.DeleteDatabase dbPath
    End If
    
    ' Step 2: Create the database file
    If Access_Database_Creation.CreateEmptyDatabase(dbPath) Then
        ' Step 3: Create all tables
        If Access_Table_Definitions.CreateAllTables(dbPath) Then
            ' Step 4: Create views (optional)
            Access_Relationships.CreateAllViews dbPath
            
            ' Step 5: Import CSV data into tables
            cntFG = Access_ImportMasterData.ImportMasterDataFinishedGoods(dbPath)
            cntAT = Access_ImportMasterData.ImportAuditTrail(dbPath)
            cntRM = Access_ImportRawMaterial.ImportRawMaterialMaster(dbPath)
            cntRMC = Access_ImportRawMaterial.ImportRMChanges(dbPath)
            
            If cntFG < 0 Or cntAT < 0 Or cntRM < 0 Or cntRMC < 0 Then
                PathResolver.ShowMessage "Failed to import CSV data.", vbCritical, "Reinitialization Failed"
                ReinitializeDatabase = False
                Exit Function
            End If
            
            importCount = cntFG + cntAT + cntRM + cntRMC
            
            ' Step 6: Test connection
            If Access_Connection_Manager.TestConnection() Then
                PathResolver.ShowMessage "Database reinitialization completed successfully! Records imported: " & importCount, vbInformation, "Reinitialization Complete"
                ReinitializeDatabase = True
                Exit Function
            Else
                PathResolver.ShowMessage "Failed to test database connection.", vbCritical, "Reinitialization Failed"
                ReinitializeDatabase = False
                Exit Function
            End If
        Else
            PathResolver.ShowMessage "Failed to create tables. Aborting reinitialization.", vbCritical, "Reinitialization Failed"
            ReinitializeDatabase = False
            Exit Function
        End If
    Else
        PathResolver.ShowMessage "Failed to create database file. Aborting reinitialization.", vbCritical, "Reinitialization Failed"
        ReinitializeDatabase = False
        Exit Function
    End If
    
ErrorHandler:
    PathResolver.ShowMessage "Error during database reinitialization: " & Err.Description, vbCritical, "Reinitialization Failed"
    ReinitializeDatabase = False
End Function

'------------------------------------------------------------------------------
' Function: GetImportSummary
' Purpose: Get a summary of the last import operation
' Returns: String with import summary
'------------------------------------------------------------------------------
Public Function GetImportSummary() As String
    On Error GoTo ErrorHandler
    
    Dim dbPath As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim tableName As Variant
    Dim totalRecords As Long
    
    dbPath = PathResolver.GetDatabasePath()
    
    If Dir(dbPath) = "" Then
        GetImportSummary = "Database does not exist."
        Exit Function
    End If
    
    Set conn = Access_Connection_Manager.GetDatabaseConnection()
    If conn Is Nothing Then
        GetImportSummary = "Failed to connect to database"
        Exit Function
    End If
    
    totalRecords = 0
    For Each tableName In Array("MasterDataFinishedGoods", "RawMaterialMaster", "AuditTrail", "RMChanges")
        sql = "SELECT COUNT(*) AS [Count] FROM [" & CStr(tableName) & "]"
        Set rs = Access_Connection_Manager.ExecuteQueryAndReturnRecordset(conn, sql)
        If Not rs Is Nothing Then
            totalRecords = totalRecords + rs!Count
            rs.Close
            Set rs = Nothing
        End If
    Next tableName
    
    conn.Close
    Set conn = Nothing
    
    GetImportSummary = "Total records in database: " & totalRecords
    Exit Function
    
ErrorHandler:
    GetImportSummary = "Error getting import summary: " & Err.Description
End Function

'------------------------------------------------------------------------------
' Function: GetDatabaseStatus
' Purpose: Get the current status of the database
' Returns: String with database status
'------------------------------------------------------------------------------
Public Function GetDatabaseStatus() As String
    On Error GoTo ErrorHandler
    
    Dim dbPath As String
    Dim conn As ADODB.Connection
    Dim tableName As Variant
    Dim tableCount As Long
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    dbPath = PathResolver.GetDatabasePath()
    
    If Dir(dbPath) = "" Then
        GetDatabaseStatus = "Database does not exist."
        Exit Function
    End If
    
    Set conn = Access_Connection_Manager.GetDatabaseConnection()
    If conn Is Nothing Then
        GetDatabaseStatus = "Database exists but connection failed."
        Exit Function
    End If
    
    tableCount = 0
    For Each tableName In Array("MasterDataFinishedGoods", "RawMaterialMaster", "AuditTrail", "RMChanges")
        sql = "SELECT COUNT(*) AS [Count] FROM [" & CStr(tableName) & "]"
        Set rs = Access_Connection_Manager.ExecuteQueryAndReturnRecordset(conn, sql)
        If Not rs Is Nothing Then
            tableCount = tableCount + 1
            rs.Close
            Set rs = Nothing
        End If
    Next tableName
    
    conn.Close
    Set conn = Nothing
    
    GetDatabaseStatus = "Database Status: OK" & vbCrLf & "Tables Created: " & tableCount & "/4" & vbCrLf & "Connection: Active"
    Exit Function
    
ErrorHandler:
    GetDatabaseStatus = "Error checking database status: " & Err.Description
End Function

'------------------------------------------------------------------------------
' Sub: ShowDatabaseInfo
' Purpose: Display detailed information about the database in a message box
'------------------------------------------------------------------------------
Public Sub ShowDatabaseInfo()
    On Error GoTo ErrorHandler
    
    Dim dbPath As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim info As String
    Dim tableName As Variant
    
    dbPath = PathResolver.GetDatabasePath()
    
    If Dir(dbPath) = "" Then
        PathResolver.ShowMessage "Database does not exist at: " & dbPath, vbExclamation, "Database Info"
        Exit Sub
    End If
    
    Set conn = Access_Connection_Manager.GetDatabaseConnection()
    If conn Is Nothing Then
        PathResolver.ShowMessage "Failed to connect to database", vbCritical, "Database Info"
        Exit Sub
    End If
    
    info = "=== Database Information ===" & vbCrLf & vbCrLf
    
    sql = "SELECT [Name] FROM [MSysObjects] WHERE [Type]=-32768 AND [Name] NOT LIKE 'MSys*'"
    Set rs = Access_Connection_Manager.ExecuteQueryAndReturnRecordset(conn, sql)
    
    If Not rs Is Nothing Then
        info = info & "Tables:" & vbCrLf
        Do While Not rs.EOF
            info = info & "  - " & rs!Name & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    
    For Each tableName In Array("MasterDataFinishedGoods", "RawMaterialMaster", "AuditTrail", "RMChanges")
        sql = "SELECT COUNT(*) AS [Count] FROM [" & CStr(tableName) & "]"
        Set rs = Access_Connection_Manager.ExecuteQueryAndReturnRecordset(conn, sql)
        If Not rs Is Nothing Then
            info = info & vbCrLf & CStr(tableName) & ": " & rs!Count & " records"
            rs.Close
            Set rs = Nothing
        End If
    Next tableName
    
    conn.Close
    Set conn = Nothing
    
    PathResolver.ShowMessage info, vbInformation, "Database Information"
    Exit Sub
    
ErrorHandler:
    PathResolver.ShowMessage "Error displaying database information: " & Err.Description, vbCritical, "Database Info"
End Sub
