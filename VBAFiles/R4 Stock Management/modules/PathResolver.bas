Attribute VB_Name = "PathResolver"
Option Explicit

' ========================================================================
' PATH RESOLVER
' Detects where this workbook lives and resolves the data folders:
'   REPO : inside the development repository (<repo>\DataFiles\...)
'   DEV  : customer ...\DEV\ tree
'   PROD : customer ...\DATA\ tree
' Layout under the data root (same for every environment):
'   <root>\Master Data\Tables\
'   <root>\Material Management\Tables\
' Override: set the QPMS_DATA_ROOT environment variable to a data root
' (used by the test runner, which opens copies outside the repository).
' ========================================================================

Private Const ENV_VAR_DATA_ROOT As String = "QPMS_DATA_ROOT"
Private Const FOLDER_DATAFILES As String = "DataFiles"
Private Const FOLDER_MASTER_DATA_NAME As String = "Master Data"
Private Const FOLDER_MATERIAL_MGMT_NAME As String = "Material Management"
Private Const FOLDER_TABLES_NAME As String = "Tables"

Private cachedWorkbookPath As String
Private cachedEnvironment As String
Private cachedRoot As String

Private Function AddSlash(ByVal p As String) As String
    If Len(p) > 0 And Right$(p, 1) <> "\" Then p = p & "\"
    AddSlash = p
End Function

Private Function DirExists(ByVal p As String) As Boolean
    On Error Resume Next
    DirExists = (Len(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function FindRepoDataRoot(ByVal startPath As String) As String
    Dim cur As String
    Dim parentPos As Long
    cur = startPath
    Do While Len(cur) > 0
        If DirExists(AddSlash(cur) & FOLDER_DATAFILES & "\" & FOLDER_MASTER_DATA_NAME) Then
            FindRepoDataRoot = AddSlash(cur) & FOLDER_DATAFILES & "\"
            Exit Function
        End If
        parentPos = InStrRev(cur, "\")
        If parentPos <= 3 Then Exit Do
        cur = Left$(cur, parentPos - 1)
    Loop
End Function

Private Function FindCustomerRoot(ByVal startPath As String, ByRef env As String) As String
    Dim upperPath As String
    Dim pos As Long
    upperPath = UCase$(AddSlash(startPath))

    pos = InStr(1, upperPath, "\DEV\", vbBinaryCompare)
    If pos > 0 Then
        env = "DEV"
        FindCustomerRoot = Left$(AddSlash(startPath), pos + 4)
        Exit Function
    End If

    pos = InStr(1, upperPath, "\DATA\", vbBinaryCompare)
    If pos > 0 Then
        env = "PROD"
        FindCustomerRoot = Left$(AddSlash(startPath), pos + 5)
    End If
End Function

Private Sub ResolveLocation()
    Dim wbPath As String
    Dim override As String
    Dim env As String
    Dim root As String

    wbPath = ThisWorkbook.Path
    If Len(cachedEnvironment) > 0 And StrComp(cachedWorkbookPath, wbPath, vbTextCompare) = 0 Then Exit Sub

    cachedWorkbookPath = wbPath
    cachedEnvironment = ""
    cachedRoot = ""

    override = Trim$(Environ$(ENV_VAR_DATA_ROOT))
    If Len(override) > 0 Then
        cachedRoot = AddSlash(override)
        cachedEnvironment = "REPO"
        Exit Sub
    End If

    If Len(wbPath) > 0 Then
        root = FindRepoDataRoot(wbPath)
        If Len(root) > 0 Then
            cachedRoot = root
            cachedEnvironment = "REPO"
            Exit Sub
        End If

        root = FindCustomerRoot(wbPath, env)
        If Len(root) > 0 Then
            cachedRoot = root
            cachedEnvironment = env
            Exit Sub
        End If
    End If

    ' Undetected location: fall back to the workbook-local Tables folder
    cachedEnvironment = "DEV"
    cachedRoot = ""
End Sub

' "REPO", "DEV" or "PROD"
Public Function GetDataEnvironment() As String
    ResolveLocation
    GetDataEnvironment = cachedEnvironment
End Function

' Drop-in replacement for the former PROD_DEV constant ("DEV" or "PROD")
Public Function PROD_DEV() As String
    If GetDataEnvironment() = "PROD" Then
        PROD_DEV = "PROD"
    Else
        PROD_DEV = "DEV"
    End If
End Function

Public Function GetEnvironmentFromPath() As String
    If GetDataEnvironment() = "PROD" Then
        GetEnvironmentFromPath = "PROD"
    Else
        GetEnvironmentFromPath = "DEV"
    End If
End Function

Public Function GetDataRoot() As String
    ResolveLocation
    GetDataRoot = cachedRoot
End Function

' <root>\Master Data\ (or the workbook folder when the location is undetected)
Public Function GetMasterDataRootPath() As String
    ResolveLocation
    If Len(cachedRoot) = 0 Then
        GetMasterDataRootPath = AddSlash(cachedWorkbookPath)
    Else
        GetMasterDataRootPath = cachedRoot & FOLDER_MASTER_DATA_NAME & "\"
    End If
End Function

' <root>\Material Management\ (or the workbook folder when the location is undetected)
Public Function GetMaterialManagementRootPath() As String
    ResolveLocation
    If Len(cachedRoot) = 0 Then
        GetMaterialManagementRootPath = AddSlash(cachedWorkbookPath)
    Else
        GetMaterialManagementRootPath = cachedRoot & FOLDER_MATERIAL_MGMT_NAME & "\"
    End If
End Function

Public Function GetMasterDataTablesPath() As String
    GetMasterDataTablesPath = GetMasterDataRootPath() & FOLDER_TABLES_NAME & "\"
End Function

Public Function GetMaterialManagementTablesPath() As String
    GetMaterialManagementTablesPath = GetMaterialManagementRootPath() & FOLDER_TABLES_NAME & "\"
End Function

Public Function GetTempFolderPath() As String
    GetTempFolderPath = GetMaterialManagementTablesPath() & "temp\"
End Function

Private Function IsMasterDataFile(ByVal fileName As String) As Boolean
    Select Case LCase$(fileName)
        Case "raw_material_master.csv", "rm_master_changes.csv", _
             "masterdatafinishedgoods.csv", "audittrailmasterdatafinishedgoods.csv"
            IsMasterDataFile = True
    End Select
End Function

' Full path for a table file; folder is chosen from the file name
Public Function GetTableFilePath(ByVal fileName As String) As String
    If Left$(fileName, 1) = "\" Then fileName = Mid$(fileName, 2)
    If IsMasterDataFile(fileName) Then
        GetTableFilePath = GetMasterDataTablesPath() & fileName
    Else
        GetTableFilePath = GetMaterialManagementTablesPath() & fileName
    End If
End Function

' Legacy-compatible helpers --------------------------------------------

Public Function GetWorkbookTablesPath() As String
    GetWorkbookTablesPath = GetMaterialManagementTablesPath()
End Function

Public Function GetWorkbookTableFilePath(ByVal fileName As String) As String
    GetWorkbookTableFilePath = GetTableFilePath(fileName)
End Function

' relativePath is relative to the Master Data folder, e.g. "Tables\Raw_Material_Master.csv"
Public Function GetMasterDataPath(Optional ByVal relativePath As String = "") As String
    Dim base As String
    base = GetMasterDataRootPath()
    If Len(relativePath) = 0 Then
        GetMasterDataPath = base
        Exit Function
    End If
    If Left$(relativePath, 1) = "\" Then relativePath = Mid$(relativePath, 2)
    GetMasterDataPath = base & relativePath
End Function
