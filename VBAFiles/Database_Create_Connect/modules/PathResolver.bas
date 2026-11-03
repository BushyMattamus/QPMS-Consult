Attribute VB_Name = "PathResolver"
'=============================================================================
' Module: PathResolver
' Purpose: Resolve database file paths and data directories for Database_Create_Connect
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

Private Const ENV_VAR_DATA_ROOT As String = "QPMS_DATA_ROOT"
Private Const ENV_VAR_TEST_MODE As String = "QPMS_TEST_MODE"
Private Const FOLDER_DATAFILES As String = "DataFiles"
Private Const FOLDER_MASTER_DATA_NAME As String = "Master Data"
Private Const FOLDER_TABLES_NAME As String = "Tables"
Private Const DB_FILENAME As String = "Database.accdb"

Private cachedWorkbookPath As String
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

Private Function FindCustomerRoot(ByVal startPath As String) As String
    Dim upperPath As String
    Dim pos As Long
    upperPath = UCase$(AddSlash(startPath))

    pos = InStr(1, upperPath, "\DEV\", vbBinaryCompare)
    If pos > 0 Then
        FindCustomerRoot = Left$(AddSlash(startPath), pos + 4)
        Exit Function
    End If

    pos = InStr(1, upperPath, "\DATA\", vbBinaryCompare)
    If pos > 0 Then
        FindCustomerRoot = Left$(AddSlash(startPath), pos + 5)
    End If
End Function

Private Sub ResolveLocation()
    Dim wbPath As String
    Dim override As String
    Dim root As String

    wbPath = ThisWorkbook.Path
    If Len(cachedRoot) > 0 And StrComp(cachedWorkbookPath, wbPath, vbTextCompare) = 0 Then Exit Sub

    cachedWorkbookPath = wbPath
    cachedRoot = ""

    override = Trim$(Environ$(ENV_VAR_DATA_ROOT))
    If Len(override) > 0 Then
        cachedRoot = AddSlash(override)
        Exit Sub
    End If

    If Len(wbPath) > 0 Then
        root = FindRepoDataRoot(wbPath)
        If Len(root) > 0 Then
            cachedRoot = root
            Exit Sub
        End If

        root = FindCustomerRoot(wbPath)
        If Len(root) > 0 Then
            cachedRoot = root
            Exit Sub
        End If
    End If

    ' Fallback to workbook directory DataFiles\
    cachedRoot = AddSlash(wbPath) & FOLDER_DATAFILES & "\"
End Sub

Public Function GetDataRoot() As String
    ResolveLocation
    GetDataRoot = cachedRoot
End Function

Public Function GetDatabasePath() As String
    GetDatabasePath = GetDataRoot() & DB_FILENAME
End Function

Public Function GetDataFilesDirectory() As String
    GetDataFilesDirectory = GetDataRoot()
End Function

Public Function GetCSVDataDirectory() As String
    GetCSVDataDirectory = GetDataRoot() & FOLDER_MASTER_DATA_NAME & "\" & FOLDER_TABLES_NAME & "\"
End Function

Public Function DatabaseExists() As Boolean
    On Error Resume Next
    DatabaseExists = (Len(Dir$(GetDatabasePath())) > 0)
    On Error GoTo 0
End Function

Public Function IsTestMode() As Boolean
    IsTestMode = (Trim$(Environ$(ENV_VAR_TEST_MODE)) = "1")
End Function

Public Sub ShowMessage(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "")
    If IsTestMode() Then Exit Sub
    MsgBox prompt, buttons, title
End Sub
