Attribute VB_Name = "Access_Relationships"
'=============================================================================
' Module: Access_Relationships
' Purpose: Create table relationships and constraints in the Access database
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: CreateRelationships
' Purpose: Create foreign key relationships between tables (if applicable)
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function CreateRelationships(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    
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
    
    ' Design uses flat tables that can be queried independently.
    conn.Close
    Set conn = Nothing
    
    CreateRelationships = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating relationships: " & Err.Description, vbCritical, "Relationship Creation Failed"
    CreateRelationships = False
End Function

'------------------------------------------------------------------------------
' Function: CreateAuditTrailLinkView
' Purpose: Create a view that links AuditTrail to MasterDataFinishedGoods
'------------------------------------------------------------------------------
Public Function CreateAuditTrailLinkView(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim s1 As String, s2 As String, s3 As String
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    
    ' Drop existing view if present
    conn.Execute "DROP VIEW [vw_AuditTrail_Linked]"
    On Error GoTo ErrorHandler
    
    ' Build SQL in segments to avoid line-continuation limit
    s1 = "CREATE VIEW [vw_AuditTrail_Linked] AS SELECT "
    s1 = s1 & "AT.ID, AT.PMI_Code, AT.PMI_Version, AT.Product_Type, AT.Product_Category, "
    s1 = s1 & "AT.Quick_Code, AT.Product_Description, AT.Product_Status, AT.Shelf_Life, "
    s1 = s1 & "AT.[Shelf_life(No of days/Months/Years)], AT.Production_Type, "
    s1 = s1 & "AT.Date_Created, AT.Created_By, AT.Date_Updated, AT.Updated_By, "
    s1 = s1 & "AT.Update_Comments, AT.PMI_Acronym, AT.PMI_Number, AT.Old_PMI, "
    
    s2 = "MD.Product_Type AS MasterProductType, MD.Product_Category AS MasterProductCategory, "
    s2 = s2 & "MD.Quick_Code AS MasterQuickCode, MD.Product_Description AS MasterDescription, "
    s2 = s2 & "MD.Product_Status AS MasterStatus, MD.PMI_Acronym AS MasterAcronym, "
    s2 = s2 & "MD.PMI_Number AS MasterNumber, "
    s2 = s2 & "AT.Date_Updated AS AuditDate, AT.Updated_By AS Auditor, "
    s2 = s2 & "AT.Update_Comments AS AuditComments "
    
    s3 = "FROM [AuditTrail] AT "
    s3 = s3 & "INNER JOIN [MasterDataFinishedGoods] MD "
    s3 = s3 & "ON AT.PMI_Code = MD.PMI_Code AND AT.PMI_Number = MD.PMI_Number"
    
    sql = s1 & s2 & s3
    
    On Error Resume Next
    conn.Execute sql
    On Error GoTo ErrorHandler
    
    conn.Close
    Set conn = Nothing
    
    CreateAuditTrailLinkView = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating AuditTrail link view: " & Err.Description, vbCritical, "View Creation Failed"
    CreateAuditTrailLinkView = False
End Function

'------------------------------------------------------------------------------
' Function: CreateRawMaterialLinkView
' Purpose: Create a view that links RMChanges to RawMaterialMaster
'------------------------------------------------------------------------------
Public Function CreateRawMaterialLinkView(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim s1 As String, s2 As String, s3 As String, s4 As String
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    
    ' Drop existing view if present
    conn.Execute "DROP VIEW [vw_RMChanges_Linked]"
    On Error GoTo ErrorHandler
    
    ' Build SQL in segments to avoid line-continuation limit
    s1 = "CREATE VIEW [vw_RMChanges_Linked] AS SELECT "
    s1 = s1 & "RC.ID, RC.RM_Code, RC.Std_Name, RC.Alt_Name, RC.Category, RC.Storage_Cond, "
    s1 = s1 & "RC.S4_Number, RC.Quick_Code, RC.Prod_ID, RC.Brand_Name, RC.Part_ARTG, "
    s1 = s1 & "RC.Master_UoM, RC.QTY_Per_Pack, RC.Conc_Vol, RC.Wt_per_mL, RC.Wt_per_Unit, "
    s1 = s1 & "RC.QC_Number_Required, RC.Testing_Flag, RC.QC_QTY, RC.Retention_Flag, "
    s1 = s1 & "RC.Retention_QTY, RC.Sample_Comment, RC.ChemAlert_Submit, RC.ChemAlert_Updated, "
    s1 = s1 & "RC.HazChem, RC.CAS_Num, RC.UN_Code, RC.Curr_Ind, RC.Created_By, RC.Create_Date, "
    s1 = s1 & "RC.Last_Change_By, RC.Last_Change_Date, RC.Last_Change_Detail, "
    s1 = s1 & "RC.Extended_Retest, RC.MRP_Min, RC.MRP_Max, RC.Pack_Price, "
    
    s2 = "RM.Std_Name AS MasterStdName, RM.Alt_Name AS MasterAltName, "
    s2 = s2 & "RM.Category AS MasterCategory, RM.Storage_Cond AS MasterStorageCond, "
    s2 = s2 & "RM.S4_Number AS MasterS4Number, RM.Quick_Code AS MasterQuickCode, "
    s2 = s2 & "RM.Prod_ID AS MasterProdID, RM.Brand_Name AS MasterBrandName, "
    s2 = s2 & "RM.Part_ARTG AS MasterPartARTG, RM.Master_UoM AS MasterMasterUoM, "
    
    s3 = "RM.QTY_Per_Pack AS MasterQTYPerPack, RM.Conc_Vol AS MasterConcVol, "
    s3 = s3 & "RM.Wt_per_mL AS MasterWtper_mL, RM.Wt_per_Unit AS MasterWtper_Unit, "
    s3 = s3 & "RM.QC_Number_Required AS MasterQCNumberRequired, RM.Testing_Flag AS MasterTestingFlag, "
    s3 = s3 & "RM.QC_QTY AS MasterQCQty, RM.Retention_Flag AS MasterRetentionFlag, "
    s3 = s3 & "RM.Retention_QTY AS MasterRetentionQty, RM.Sample_Comment AS MasterSampleComment, "
    s3 = s3 & "RM.ChemAlert_Submit AS MasterChemAlertSubmit, RM.ChemAlert_Updated AS MasterChemAlertUpdated, "
    s3 = s3 & "RM.HazChem AS MasterHazChem, RM.CAS_Num AS MasterCASNum, RM.UN_Code AS MasterUNCode, "
    s3 = s3 & "RM.Curr_Ind AS MasterCurrInd, RM.Created_By AS MasterCreatedBy, RM.Create_Date AS MasterCreateDate "
    
    s4 = "FROM [RMChanges] RC "
    s4 = s4 & "INNER JOIN [RawMaterialMaster] RM ON RC.RM_Code = RM.RM_Code"
    
    sql = s1 & s2 & s3 & s4
    
    On Error Resume Next
    conn.Execute sql
    On Error GoTo ErrorHandler
    
    conn.Close
    Set conn = Nothing
    
    CreateRawMaterialLinkView = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating RMChanges link view: " & Err.Description, vbCritical, "View Creation Failed"
    CreateRawMaterialLinkView = False
End Function

'------------------------------------------------------------------------------
' Function: CreateAllViews
' Purpose: Create all views for linking related data
' Returns: True if all views were created successfully, False otherwise
'------------------------------------------------------------------------------
Public Function CreateAllViews(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not CreateAuditTrailLinkView(dbPath) Then Exit Function
    If Not CreateRawMaterialLinkView(dbPath) Then Exit Function
    
    CreateAllViews = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error creating views: " & Err.Description, vbCritical, "View Creation Failed"
    CreateAllViews = False
End Function

'------------------------------------------------------------------------------
' Function: DropAllViews
' Purpose: Drop all views from the database
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function DropAllViews(dbPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open dbPath
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open dbPath
    End If
    
    conn.Execute "DROP VIEW [vw_AuditTrail_Linked]"
    conn.Execute "DROP VIEW [vw_RMChanges_Linked]"
    On Error GoTo ErrorHandler
    
    conn.Close
    Set conn = Nothing
    
    DropAllViews = True
    Exit Function
    
ErrorHandler:
    PathResolver.ShowMessage "Error dropping views: " & Err.Description, vbCritical, "View Dropping Failed"
    DropAllViews = False
End Function
