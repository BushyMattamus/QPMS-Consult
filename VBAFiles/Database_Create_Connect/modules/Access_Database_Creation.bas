Attribute VB_Name = "Access_Database_Creation"
'=============================================================================
' Module: Access_Database_Creation
' Purpose: Create new Access database file structure using ADOX
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: CreateDatabase
' Returns: True if database was created successfully or already exists, False otherwise
'------------------------------------------------------------------------------
Public Function CreateDatabase(dbPath As String) As Boolean
    CreateDatabase = CreateEmptyDatabase(dbPath)
End Function

'------------------------------------------------------------------------------
' Function: InitializeDatabaseProperties
' Purpose: Set default properties on the Access database (optional)
'------------------------------------------------------------------------------
Public Sub InitializeDatabaseProperties(dbPath As String)
    ' Reserved for database-level property initialization if needed
End Sub

'------------------------------------------------------------------------------
' Function: CreateEmptyDatabase
' Returns: True if database exists or was created, False on failure
'------------------------------------------------------------------------------
Public Function CreateEmptyDatabase(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if database already exists
    If Dir(dbPath) <> "" Then
        CreateEmptyDatabase = True
        Exit Function
    End If
    
    ' Create the database file using ADOX
    Dim cat As Object
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    On Error Resume Next
    cat.ActiveConnection.Close
    On Error GoTo ErrorHandler
    Set cat = Nothing
    
    ' Verify the file was created
    If Dir(dbPath) = "" Then
        PathResolver.ShowMessage "Failed to create database file at: " & dbPath, vbCritical, "Database Creation Failed"
        CreateEmptyDatabase = False
        Exit Function
    End If
    
    CreateEmptyDatabase = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating database: " & Err.Description, vbCritical, "Database Creation Failed"
    CreateEmptyDatabase = False
End Function

'------------------------------------------------------------------------------
' Function: DeleteDatabase
' Purpose: Delete the Access database file
' Returns: True if deletion was successful, False otherwise
'------------------------------------------------------------------------------
Public Function DeleteDatabase(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if file exists
    If Dir(dbPath) = "" Then
        DeleteDatabase = True  ' Nothing to delete
        Exit Function
    End If
    
    ' Delete the file
    Kill dbPath
    
    ' Verify deletion
    If Dir(dbPath) <> "" Then
        PathResolver.ShowMessage "Failed to delete database: " & dbPath, vbCritical, "Deletion Failed"
        DeleteDatabase = False
        Exit Function
    End If
    
    DeleteDatabase = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error deleting database: " & Err.Description, vbCritical, "Deletion Failed"
    DeleteDatabase = False
End Function
