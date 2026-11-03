Attribute VB_Name = "InventoryUtils"
' check this has been loaded and tested
Option Explicit

' Function to get the next sequential Transaction ID for RM_Trans.csv
Public Function GetNextTransactionID() As String
    Dim conn As Object
    Dim rs As Object
    Dim filePath As String
    Dim sql As String
    Dim maxTransID As Long
    Dim transIDStr As String
    Dim fileName As String
    
    On Error GoTo ErrHandler
    
    ' --- Initialize defaults ---
    GetNextTransactionID = "I" & Format(TRANS_ID_START_VALUE, TRANS_ID_FORMAT_STRING)
    maxTransID = 0
    
    ' --- Set file path ---
    filePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    ' --- Check if file exists ---
    If Not FileExists(filePath) Then
        GoTo CleanExit
    End If
    
    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(Left(filePath, InStrRev(filePath, "\") - 1))
    If conn Is Nothing Then
        MsgBox "Failed to open ADO connection.", vbCritical, "GetNextTransactionID"
        GoTo CleanExit
    End If
    
    ' --- SQL query to get the last Trans_ID ---
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    sql = "SELECT TOP 1 [" & RMT_HEADER_TRANS_ID & "] FROM [" & fileName & "] " & _
          "ORDER BY [" & RMT_HEADER_TRANS_ID & "] DESC"
    
    ' --- Execute query ---
    Set rs = conn.Execute(sql)
    
    ' --- Get the last Trans_ID ---
    If Not rs.EOF Then
        transIDStr = rs.Fields(RMT_HEADER_TRANS_ID).value
        
        ' Remove "I" prefix if present before converting to numeric
        If Left$(transIDStr, 1) = "I" Then
            transIDStr = Mid$(transIDStr, 2)
        End If
        
        ' Convert to numeric and calculate next ID
        If IsNumeric(transIDStr) Then
            maxTransID = CLng(transIDStr)
            GetNextTransactionID = "I" & Format(maxTransID + 1, TRANS_ID_FORMAT_STRING)
        End If
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetNextTransactionID"
    GetNextTransactionID = "I" & Format(TRANS_ID_START_VALUE, TRANS_ID_FORMAT_STRING)

CleanExit:
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If Not conn.State = 0 Then conn.Close
        Set conn = Nothing
    End If
End Function

' Function to get the next sequential QC Transaction ID for QC_Number_Changes.csv
Public Function GetNextQCTransactionID() As String
    Dim conn As Object
    Dim rs As Object
    Dim filePath As String
    Dim sql As String
    Dim maxTransID As Long
    Dim transIDStr As String

    On Error GoTo ErrorHandler

    ' --- Initialize defaults ---
    GetNextQCTransactionID = "Q" & Format(TRANS_ID_START_VALUE, TRANS_ID_FORMAT_STRING)
    maxTransID = 0

    ' --- Set file path ---
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)

    ' --- Check if file exists ---
    If Not FileExists(filePath) Then
        GoTo CleanExit
    End If

    ' --- Open ADO connection using the new function ---
    Set conn = OpenADOConnection(Left(filePath, InStrRev(filePath, "\") - 1))
    If conn Is Nothing Then
        MsgBox "Failed to open ADO connection.", vbCritical, "GetNextQCTransactionID"
        GoTo CleanExit
    End If

    ' --- SQL query to get the last QC_Trans_ID ---
    sql = "SELECT TOP 1 [" & RMT_HEADER_QC_TRANS_ID & "] FROM [" & Mid(filePath, InStrRev(filePath, "\") + 1) & "] " & _
          "ORDER BY [" & RMT_HEADER_QC_TRANS_ID & "] DESC"

    ' --- Execute query ---
    Set rs = conn.Execute(sql)

    ' --- Get the last QC_Trans_ID ---
    If Not rs.EOF Then
        transIDStr = rs.Fields(RMT_HEADER_QC_TRANS_ID).value

        ' Remove "Q" prefix if present before converting to numeric
        If Left$(transIDStr, 1) = "Q" Then
            transIDStr = Mid$(transIDStr, 2)
        End If

        ' Convert to numeric and calculate next ID
        If IsNumeric(transIDStr) Then
            maxTransID = CLng(transIDStr)
            GetNextQCTransactionID = "Q" & Format(maxTransID + 1, TRANS_ID_FORMAT_STRING)
        End If
    End If

    GoTo CleanExit

ErrorHandler:
    ' On error, return default starting ID
    GetNextQCTransactionID = "Q" & Format(TRANS_ID_START_VALUE, TRANS_ID_FORMAT_STRING)

CleanExit:
    ' --- Clean up ---
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If Not conn.State = 0 Then conn.Close
        Set conn = Nothing
    End If
End Function

' Get the balance of a specific transaction type for a QC Number from RM_Trans.csv
Public Function GetQCTransactionBalance(ByVal qcNumber As String, ByVal transType As String) As Double
    On Error GoTo ErrorHandler

    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    Dim balance As Double

    GetQCTransactionBalance = 0
    balance = 0

    If Trim$(qcNumber) = "" Or Trim$(transType) = "" Then
        GoTo CleanExit
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempPath = "" Then
        GoTo CleanExit
    End If
    
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    Set conn = OpenADOConnection(folderPath)

    If conn Is Nothing Then
        GoTo CleanExit
    End If

    sql = "SELECT SUM([" & RMT_HEADER_QUANTITY & "]) AS Balance " & _
          "FROM [" & fileName & "] " & _
          "WHERE [" & RMT_HEADER_QC_NUMBER & "] = '" & Replace(qcNumber, "'", "''") & "' " & _
          "AND [" & RMT_HEADER_TRANS_TYPE & "] = '" & Replace(transType, "'", "''") & "'"

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        If Not IsNull(rs.Fields("Balance").value) Then
            balance = CDbl(rs.Fields("Balance").value)
        End If
    End If

    GetQCTransactionBalance = balance
    GoTo CleanExit

ErrorHandler:
    GetQCTransactionBalance = 0
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Function
End Function

' New subroutine to load QC balances from RM_Trans.csv as an array
Public Sub LoadQCBalancesFromRMTrans(ByRef balanceLookup As Object)
    On Error GoTo LoadBalanceError

    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    Dim qcNumber As String
    Dim balance As Double

    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempPath = "" Then
        GoTo LoadBalanceCleanup
    End If
    
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    Set conn = OpenADOConnection(folderPath)

    If conn Is Nothing Then
        GoTo LoadBalanceCleanup
    End If

    sql = "SELECT [" & RMT_HEADER_QC_NUMBER & "], " & _
          "SUM([" & RMT_HEADER_QUANTITY & "]) AS Balance " & _
          "FROM [" & fileName & "] " & _
          "GROUP BY [" & RMT_HEADER_QC_NUMBER & "]"

    Set rs = conn.Execute(sql)

    Do While Not rs.EOF
        qcNumber = Trim$(rs.Fields(RMT_HEADER_QC_NUMBER).value & "")
        If qcNumber <> "" Then
            If Not IsNull(rs.Fields("Balance").value) Then
                balance = CDbl(rs.Fields("Balance").value)
                balanceLookup(qcNumber) = balance
            End If
        End If
        rs.MoveNext
    Loop

    GoTo LoadBalanceCleanup

LoadBalanceError:

LoadBalanceCleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
End Sub

' Load QC balances and total sample quantities from RM_Trans.csv
Public Sub LoadQCBalancesAndSamplesFromRMTrans(ByRef balanceLookup As Object, ByRef sampleLookup As Object)
    On Error GoTo LoadBalanceError

    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    Dim qcNumber As String
    Dim balance As Double
    Dim sampleBalance As Double

    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempPath = "" Then
        GoTo LoadBalanceCleanup
    End If
    
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    Set conn = OpenADOConnection(folderPath)

    If conn Is Nothing Then
        GoTo LoadBalanceCleanup
    End If

    sql = "SELECT [" & RMT_HEADER_QC_NUMBER & "], " & _
          "SUM([" & RMT_HEADER_QUANTITY & "]) AS Balance " & _
          "FROM [" & fileName & "] " & _
          "GROUP BY [" & RMT_HEADER_QC_NUMBER & "]"

    Set rs = conn.Execute(sql)

    Do While Not rs.EOF
        qcNumber = Trim$(rs.Fields(RMT_HEADER_QC_NUMBER).value & "")
        If qcNumber <> "" Then
            If Not IsNull(rs.Fields("Balance").value) Then
                balance = CDbl(rs.Fields("Balance").value)
                balanceLookup(qcNumber) = balance
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    sql = "SELECT [" & RMT_HEADER_QC_NUMBER & "], " & _
          "SUM([" & RMT_HEADER_QUANTITY & "]) AS SampleBalance " & _
          "FROM [" & fileName & "] " & _
          "WHERE [" & RMT_HEADER_TRANS_TYPE & "] IN ('" & TRANS_TYPE_SAMPLE & "', '" & TRANS_TYPE_SAMPLING_REVERSAL & "') " & _
          "GROUP BY [" & RMT_HEADER_QC_NUMBER & "]"

    Set rs = conn.Execute(sql)

    Do While Not rs.EOF
        qcNumber = Trim$(rs.Fields(RMT_HEADER_QC_NUMBER).value & "")
        If qcNumber <> "" Then
            If Not IsNull(rs.Fields("SampleBalance").value) Then
                sampleBalance = CDbl(rs.Fields("SampleBalance").value)
                sampleLookup(qcNumber) = sampleBalance
            End If
        End If
        rs.MoveNext
    Loop

    GoTo LoadBalanceCleanup

LoadBalanceError:

LoadBalanceCleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
End Sub


' Returns next PI number in format "PI" + TRANS_ID_FORMAT_STRING (e.g. "PI000001")
Public Function GetNextPINumber() As String
    Dim fso As Object, ts As Object
    Dim filePath As String, fileContent As String
    Dim lines() As String, i As Long
    Dim header As String, parts() As String
    Dim maxNum As Long, candidate As String, numericPart As Long
    Dim prefix As String
    
    On Error GoTo ErrHandler
    prefix = "PI"
    maxNum = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_PHYSICAL_INVENTORY)
    
    If Not fso.FileExists(filePath) Then
        GetNextPINumber = prefix & Format$(TRANS_ID_START_VALUE, TRANS_ID_FORMAT_STRING)
        GoTo CleanExit
    End If
    
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    
    If Trim$(fileContent) = "" Then
        GetNextPINumber = prefix & Format$(TRANS_ID_START_VALUE, TRANS_ID_FORMAT_STRING)
        GoTo CleanExit
    End If
    
    lines = Split(fileContent, vbCrLf)
    
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            parts = Split(lines(i), ",")
            candidate = Replace(Replace(parts(0), """", ""), prefix, "")
            If IsNumeric(candidate) Then
                numericPart = CLng(candidate)
                If numericPart > maxNum Then maxNum = numericPart
            End If
        End If
    Next i
    
    GetNextPINumber = prefix & Format$(maxNum + 1, TRANS_ID_FORMAT_STRING)
    GoTo CleanExit

ErrHandler:
    GetNextPINumber = prefix & Format$(TRANS_ID_START_VALUE, TRANS_ID_FORMAT_STRING)
CleanExit:
    On Error Resume Next
    Set ts = Nothing
    Set fso = Nothing
End Function

' ========================================================================
' Get QC Balance from RM_Trans.csv using ADO connection
' Sums all Quantity values for the specified QC Number
' ========================================================================
Public Function GetQCBalanceADO(ByVal qcNumber As String) As Double
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    Dim balance As Double
    
    ' --- Initialize return value ---
    GetQCBalanceADO = 0
    balance = 0
    
    ' --- Validate input ---
    If Trim$(qcNumber) = "" Then
        GoTo CleanExit
    End If
    
    ' --- Create FSO ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' --- Create temporary file copy ---
    tempPath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempPath = "" Then
        GoTo CleanExit
    End If
    
    ' --- Extract folder and file name ---
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(folderPath)
    
    If conn Is Nothing Then
        GoTo CleanExit
    End If
    
    ' --- Build and execute SQL query ---
    sql = "SELECT SUM([" & RMT_HEADER_QUANTITY & "]) AS TotalQty " & _
          "FROM [" & fileName & "] " & _
          "WHERE [" & RMT_HEADER_QC_NUMBER & "] = '" & qcNumber & "'"
    
    Set rs = conn.Execute(sql)
    
    ' --- Get result ---
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("TotalQty").value) Then
            balance = CDbl(rs.Fields("TotalQty").value)
        End If
    End If
    
    GetQCBalanceADO = balance
    GoTo CleanExit

' --- Error handler ---
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetQCBalanceADO"
    GoTo CleanExit

' --- Clean exit & object release ---
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Function
End Function
