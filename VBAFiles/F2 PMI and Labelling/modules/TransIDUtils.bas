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

' --- Get the next transaction ID for Finished Goods Batch Audit ---
Public Function GetNextFinishedGoodsBatchMaster() As String
    Dim conn As Object
    Dim rs As Object
    Dim folderPath As String
    Dim masterFile As String
    Dim sql As String
    Dim output As Long
    Dim nextNumber As Long
    Dim batchNumber As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    ' --- Get folder path and master file ---
    folderPath = GetWorkbookTablesPath()
    masterFile = FILE_FINISHED_GOODS_BATCH_MASTER
    
    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(folderPath)
    If conn Is Nothing Then
        MsgBox "Failed to open connection to Tables folder", vbCritical, "GetNextFinishedGoodsBatchMaster"
        GoTo CleanExit
    End If
    
    ' --- Query for count of records ---
    sql = "SELECT COUNT(*) AS RecordCount FROM [" & masterFile & "]"
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        output = rs.Fields("RecordCount").value
        nextNumber = output + 1
    Else
        nextNumber = 1
    End If
    
    Debug.Print "Total Records: " & output
    Debug.Print "Next Number: " & nextNumber
    Debug.Print "---"
    
    ' --- Query all batch numbers ---
    sql = "SELECT Batch_Number FROM [" & masterFile & "] ORDER BY Batch_Number"
    Set rs = conn.Execute(sql)
    
    i = 1
    Do While Not rs.EOF
        batchNumber = rs.Fields("Batch_Number").value
        Debug.Print i & ": " & batchNumber
        rs.MoveNext
        i = i + 1
    Loop
    
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetNextFinishedGoodsBatchMaster"
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Function

' --- Generate next Label Print ID in format LP000001 ---
Public Function GetNextLabelPrintID() As String
    On Error GoTo ErrHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim folderPath As String
    Dim sql As String
    Dim nextNumber As Long
    
    GetNextLabelPrintID = "LP000001"
    
    folderPath = GetWorkbookTablesPath()
    
    Set conn = OpenADOConnection(folderPath)
    If conn Is Nothing Then
        MsgBox "Failed to open connection to Tables folder", vbCritical, "GetNextLabelPrintID"
        GoTo CleanExit
    End If
    
    sql = "SELECT Print_ID FROM [" & FILE_LABEL_DETAILS & "] " & _
          "WHERE LEFT(Print_ID, 2) = 'LP' " & _
          "ORDER BY Print_ID DESC"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF And Not IsNull(rs.Fields("Print_ID").value) Then
        nextNumber = CLng(Mid$(rs.Fields("Print_ID").value, 3)) + 1
    Else
        nextNumber = 1
    End If
    
    GetNextLabelPrintID = "LP" & Format$(nextNumber, "000000")
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetNextLabelPrintID"
    GetNextLabelPrintID = "LP000001"
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function
End Function