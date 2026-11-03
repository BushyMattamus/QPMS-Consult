Attribute VB_Name = "Access_Table_Definitions"
'=============================================================================
' Module: Access_Table_Definitions
' Purpose: Define table structures with fields and data types for Access database
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: CreateMasterDataFinishedGoodsTable
' Purpose: Create the MasterDataFinishedGoods table with all fields
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function CreateMasterDataFinishedGoodsTable(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim tableName As String
    tableName = "MasterDataFinishedGoods"
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    
    ' Drop existing table if present
    conn.Execute "DROP TABLE [" & tableName & "]"
    On Error GoTo ErrorHandler
    
    ' Build CREATE TABLE in segments to avoid line-continuation limit
    sql = "CREATE TABLE [" & tableName & "] ("
    sql = sql & "[ID] AUTOINCREMENT PRIMARY KEY, "
    sql = sql & "[PMI_Code] Text(50) NOT NULL, "
    sql = sql & "[PMI_Version] Number, "
    sql = sql & "[Product_Type] Text(50), "
    sql = sql & "[Product_Category] Text(50), "
    sql = sql & "[Quick_Code] Text(50), "
    sql = sql & "[Product_Description] Text(255), "
    sql = sql & "[Product_Status] Text(50), "
    sql = sql & "[Shelf_Life] Text(50), "
    sql = sql & "[Shelf_life(No of days/Months/Years)] Text(50), "
    sql = sql & "[Production_Type] Text(50), "
    sql = sql & "[Date_Created] DateTime, "
    sql = sql & "[Created_By] Text(100), "
    sql = sql & "[Date_Updated] DateTime, "
    sql = sql & "[Updated_By] Text(100), "
    sql = sql & "[Update_Comments] Text(255), "
    sql = sql & "[PMI_Acronym] Text(50), "
    sql = sql & "[PMI_Number] Number, "
    sql = sql & "[Old_PMI] Text(50)"
    sql = sql & ")"
    
    conn.Execute sql
    
    ' Create index on PMI_Code for faster lookups
    sql = "CREATE INDEX [IX_MasterDataFinishedGoods_PMI_Code] ON [" & tableName & "] ([PMI_Code])"
    conn.Execute sql
    
    conn.Close
    Set conn = Nothing
    
    CreateMasterDataFinishedGoodsTable = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating MasterDataFinishedGoods table: " & Err.Description, vbCritical, "Table Creation Failed"
    CreateMasterDataFinishedGoodsTable = False
End Function

'------------------------------------------------------------------------------
' Function: CreateRawMaterialMasterTable
' Purpose: Create the RawMaterialMaster table with all fields
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function CreateRawMaterialMasterTable(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim tableName As String
    tableName = "RawMaterialMaster"
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    
    ' Drop existing table if present
    conn.Execute "DROP TABLE [" & tableName & "]"
    On Error GoTo ErrorHandler
    
    ' Build CREATE TABLE in segments to avoid line-continuation limit
    sql = "CREATE TABLE [" & tableName & "] ("
    sql = sql & "[ID] AUTOINCREMENT PRIMARY KEY, "
    sql = sql & "[RM_Code] Text(50) NOT NULL, "
    sql = sql & "[Std_Name] Text(255), "
    sql = sql & "[Alt_Name] Text(255), "
    sql = sql & "[Category] Text(50), "
    sql = sql & "[Storage_Cond] Text(50), "
    sql = sql & "[S4_Number] Text(50), "
    sql = sql & "[Quick_Code] Text(50), "
    sql = sql & "[Prod_ID] Text(50), "
    sql = sql & "[Brand_Name] Text(255), "
    sql = sql & "[Part_ARTG] Text(100), "
    sql = sql & "[Master_UoM] Text(50), "
    sql = sql & "[QTY_Per_Pack] Number, "
    sql = sql & "[Conc_Vol] Text(50), "
    sql = sql & "[Wt_per_mL] Number, "
    sql = sql & "[Wt_per_Unit] Number, "
    sql = sql & "[QC_Number_Required] YesNo, "
    sql = sql & "[Testing_Flag] YesNo, "
    sql = sql & "[QC_QTY] Number, "
    sql = sql & "[Retention_Flag] YesNo, "
    sql = sql & "[Retention_QTY] Number, "
    sql = sql & "[Sample_Comment] Text(255), "
    sql = sql & "[ChemAlert_Submit] YesNo, "
    sql = sql & "[ChemAlert_Updated] YesNo, "
    sql = sql & "[HazChem] YesNo, "
    sql = sql & "[CAS_Num] Text(100), "
    sql = sql & "[UN_Code] Text(50), "
    sql = sql & "[Curr_Ind] YesNo, "
    sql = sql & "[Created_By] Text(100), "
    sql = sql & "[Create_Date] DateTime, "
    sql = sql & "[Last_Change_By] Text(100), "
    sql = sql & "[Last_Change_Date] DateTime, "
    sql = sql & "[Last_Change_Detail] Text(255), "
    sql = sql & "[Extended_Retest] YesNo, "
    sql = sql & "[MRP_Min] Number, "
    sql = sql & "[MRP_Max] Number, "
    sql = sql & "[Pack_Price] Currency"
    sql = sql & ")"
    
    conn.Execute sql
    
    ' Create index on RM_Code for faster lookups
    sql = "CREATE INDEX [IX_RawMaterialMaster_RM_Code] ON [" & tableName & "] ([RM_Code])"
    conn.Execute sql
    
    conn.Close
    Set conn = Nothing
    
    CreateRawMaterialMasterTable = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating RawMaterialMaster table: " & Err.Description, vbCritical, "Table Creation Failed"
    CreateRawMaterialMasterTable = False
End Function

'------------------------------------------------------------------------------
' Function: CreateAuditTrailTable
' Purpose: Create the AuditTrail table with all fields
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function CreateAuditTrailTable(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim tableName As String
    tableName = "AuditTrail"
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    
    ' Drop existing table if present
    conn.Execute "DROP TABLE [" & tableName & "]"
    On Error GoTo ErrorHandler
    
    ' Build CREATE TABLE in segments to avoid line-continuation limit
    sql = "CREATE TABLE [" & tableName & "] ("
    sql = sql & "[ID] AUTOINCREMENT PRIMARY KEY, "
    sql = sql & "[PMI_Code] Text(50), "
    sql = sql & "[PMI_Version] Number, "
    sql = sql & "[Product_Type] Text(50), "
    sql = sql & "[Product_Category] Text(50), "
    sql = sql & "[Quick_Code] Text(50), "
    sql = sql & "[Product_Description] Text(255), "
    sql = sql & "[Product_Status] Text(50), "
    sql = sql & "[Shelf_Life] Text(50), "
    sql = sql & "[Shelf_life(No of days/Months/Years)] Text(50), "
    sql = sql & "[Production_Type] Text(50), "
    sql = sql & "[Date_Created] DateTime, "
    sql = sql & "[Created_By] Text(100), "
    sql = sql & "[Date_Updated] DateTime, "
    sql = sql & "[Updated_By] Text(100), "
    sql = sql & "[Update_Comments] Text(255), "
    sql = sql & "[PMI_Acronym] Text(50), "
    sql = sql & "[PMI_Number] Number, "
    sql = sql & "[Old_PMI] Text(50)"
    sql = sql & ")"
    
    conn.Execute sql
    
    ' Create index on PMI_Code for faster lookups
    sql = "CREATE INDEX [IX_AuditTrail_PMI_Code] ON [" & tableName & "] ([PMI_Code])"
    conn.Execute sql
    
    conn.Close
    Set conn = Nothing
    
    CreateAuditTrailTable = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating AuditTrail table: " & Err.Description, vbCritical, "Table Creation Failed"
    CreateAuditTrailTable = False
End Function

'------------------------------------------------------------------------------
' Function: CreateRMChangesTable
' Purpose: Create the RMChanges table (same structure as RawMaterialMaster)
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function CreateRMChangesTable(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim tableName As String
    tableName = "RMChanges"
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    
    ' Drop existing table if present
    conn.Execute "DROP TABLE [" & tableName & "]"
    On Error GoTo ErrorHandler
    
    ' Build CREATE TABLE in segments to avoid line-continuation limit
    sql = "CREATE TABLE [" & tableName & "] ("
    sql = sql & "[ID] AUTOINCREMENT PRIMARY KEY, "
    sql = sql & "[RM_Code] Text(50) NOT NULL, "
    sql = sql & "[Std_Name] Text(255), "
    sql = sql & "[Alt_Name] Text(255), "
    sql = sql & "[Category] Text(50), "
    sql = sql & "[Storage_Cond] Text(50), "
    sql = sql & "[S4_Number] Text(50), "
    sql = sql & "[Quick_Code] Text(50), "
    sql = sql & "[Prod_ID] Text(50), "
    sql = sql & "[Brand_Name] Text(255), "
    sql = sql & "[Part_ARTG] Text(100), "
    sql = sql & "[Master_UoM] Text(50), "
    sql = sql & "[QTY_Per_Pack] Number, "
    sql = sql & "[Conc_Vol] Text(50), "
    sql = sql & "[Wt_per_mL] Number, "
    sql = sql & "[Wt_per_Unit] Number, "
    sql = sql & "[QC_Number_Required] YesNo, "
    sql = sql & "[Testing_Flag] YesNo, "
    sql = sql & "[QC_QTY] Number, "
    sql = sql & "[Retention_Flag] YesNo, "
    sql = sql & "[Retention_QTY] Number, "
    sql = sql & "[Sample_Comment] Text(255), "
    sql = sql & "[ChemAlert_Submit] YesNo, "
    sql = sql & "[ChemAlert_Updated] YesNo, "
    sql = sql & "[HazChem] YesNo, "
    sql = sql & "[CAS_Num] Text(100), "
    sql = sql & "[UN_Code] Text(50), "
    sql = sql & "[Curr_Ind] YesNo, "
    sql = sql & "[Created_By] Text(100), "
    sql = sql & "[Create_Date] DateTime, "
    sql = sql & "[Last_Change_By] Text(100), "
    sql = sql & "[Last_Change_Date] DateTime, "
    sql = sql & "[Last_Change_Detail] Text(255), "
    sql = sql & "[Extended_Retest] YesNo, "
    sql = sql & "[MRP_Min] Number, "
    sql = sql & "[MRP_Max] Number, "
    sql = sql & "[Pack_Price] Currency"
    sql = sql & ")"
    
    conn.Execute sql
    
    ' Create index on RM_Code for faster lookups
    sql = "CREATE INDEX [IX_RMChanges_RM_Code] ON [" & tableName & "] ([RM_Code])"
    conn.Execute sql
    
    conn.Close
    Set conn = Nothing
    
    CreateRMChangesTable = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating RMChanges table: " & Err.Description, vbCritical, "Table Creation Failed"
    CreateRMChangesTable = False
End Function

'------------------------------------------------------------------------------
' Function: CreateAllTables
' Purpose: Create all tables in the database
' Returns: True if all tables were created successfully, False otherwise
'------------------------------------------------------------------------------
Public Function CreateAllTables(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not CreateMasterDataFinishedGoodsTable(dbPath) Then Exit Function
    If Not CreateRawMaterialMasterTable(dbPath) Then Exit Function
    If Not CreateAuditTrailTable(dbPath) Then Exit Function
    If Not CreateRMChangesTable(dbPath) Then Exit Function
    
    CreateAllTables = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating tables: " & Err.Description, vbCritical, "Table Creation Failed"
    CreateAllTables = False
End Function
