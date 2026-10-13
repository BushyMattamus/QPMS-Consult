Attribute VB_Name = "TransIDUtils"
Option Explicit

' --- Get the next transaction ID for Finished Goods Batch Audit ---
Public Function GetNextFinishedGoodsBatchTransID() As String
    Dim conn As Object
    Dim rs As Object
    Dim folderPath As String
    Dim auditFile As String
    Dim sql As String
    Dim lastTransID As String
    Dim nextNumber As Long
    
    On Error GoTo ErrHandler
    
    ' --- Get folder path and audit file ---
    folderPath = GetWorkbookTablesPath()
    auditFile = FILE_FINISHED_GOODS_BATCH_AUDIT
    
    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(folderPath)
    If conn Is Nothing Then
        MsgBox "Failed to open connection to Tables folder", vbCritical, "GetNextFinishedGoodsBatchTransID"
        GoTo CleanExit
    End If
    
    ' --- Query for maximum Trans_Id ---
    sql = "SELECT MAX(Trans_Id) AS MaxTransID FROM [" & auditFile & "]"
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        lastTransID = rs.Fields("MaxTransID").value
        If IsNull(lastTransID) Or lastTransID = "" Then
            ' --- No records yet, start with 1 ---
            nextNumber = 1
        Else
            ' --- Extract numeric part and increment ---
            nextNumber = CLng(Mid$(lastTransID, 2)) + 1
        End If
    Else
        ' --- No records, start with 1 ---
        nextNumber = 1
    End If
    
    ' --- Format with "B" prefix ---
    GetNextFinishedGoodsBatchTransID = "B" & Format$(nextNumber, TRANS_ID_FORMAT_STRING)
    
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetNextFinishedGoodsBatchTransID"
    GetNextFinishedGoodsBatchTransID = "B000001"
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Function