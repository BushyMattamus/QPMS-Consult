Attribute VB_Name = "DropdownUtils"
'-------------------------------------------------------------------------------
' PopulateStandardNameDropdownForBatch
'-------------------------------------------------------------------------------
' Purpose: Populate Standard Name dropdown filtered by Batch Number
' Reads RM_Trans.csv to get QC Numbers for batch, then looks up Standard Names
' from QC_Numbers worksheet
'-------------------------------------------------------------------------------
Public Sub PopulateStandardNameDropdownForBatch(ByVal batchNumber As String, ByVal targetCell As Range)
    Dim conn As Object
    Dim rs As Object
    Dim fso As Object
    Dim tempFilePath As String
    Dim tempFolder As String
    Dim qcDict As Object
    Dim stdNameDict As Object
    Dim wsQC As Worksheet
    Dim colQCNumber As Long
    Dim colStdName As Long
    Dim lastRow As Long
    Dim r As Long
    Dim qcNumber As String
    Dim stdName As String
    Dim dropdownList As String
    Dim qcKeys As Variant
    Dim stdNames As Variant
    Dim qcVal As String
    Dim sqlQuery As String
    Dim j As Long
    
    On Error GoTo ErrorHandler
    
    Set qcDict = CreateObject("Scripting.Dictionary")
    Set stdNameDict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    tempFilePath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempFilePath = "" Then
        MsgBox "Unable to create temporary copy of RM_Trans.csv.", vbExclamation, "PopulateStandardNameDropdownForBatch"
        GoTo CleanExit
    End If
    
    tempFolder = fso.GetParentFolderName(tempFilePath)
    
    Set conn = OpenADOConnection(tempFolder)
    
    If conn Is Nothing Then
        MsgBox "Unable to establish database connection.", vbCritical, "PopulateStandardNameDropdownForBatch"
        GoTo CleanExit
    End If
    
    sqlQuery = "SELECT DISTINCT " & RMT_HEADER_QC_NUMBER & " " & _
               "FROM [" & fso.GetFileName(tempFilePath) & "] " & _
               "WHERE " & RMT_HEADER_FG_BATCH & " = '" & Replace(batchNumber, "'", "''") & "' " & _
               "AND " & RMT_HEADER_TRANS_TYPE & " = '" & TRANS_TYPE_CONSUMED & "' " & _
               "AND (" & RMT_HEADER_REVERSED_FLAG & " IS NULL OR " & RMT_HEADER_REVERSED_FLAG & " <> 'Y')"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            qcVal = Trim$(GetFieldValue(rs, RMT_HEADER_QC_NUMBER) & "")
            If qcVal <> "" And Not qcDict.Exists(qcVal) Then
                qcDict.Add qcVal, True
            End If
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    If qcDict.Count = 0 Then
        ClearDropdownValidation targetCell
        GoTo CleanExit
    End If
    
    ' --- Lookup Standard Names from QC_Number worksheet ---
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colStdName = FindHeaderCol(wsQC, 1, HEADER_STD_NAME)
    
    If colQCNumber = 0 Or colStdName = 0 Then
        MsgBox "Required columns not found in QC_Number worksheet.", vbCritical, "PopulateStandardNameDropdownForBatch"
        GoTo CleanExit
    End If
    
    lastRow = wsQC.Cells(wsQC.Rows.Count, colQCNumber).End(xlUp).Row
    
    qcKeys = qcDict.Keys
    
    For r = 2 To lastRow
        qcNumber = Trim$(CStr(wsQC.Cells(r, colQCNumber).value))
        
        For j = LBound(qcKeys) To UBound(qcKeys)
            If LCase$(qcNumber) = LCase$(qcKeys(j)) Then
                stdName = Trim$(CStr(wsQC.Cells(r, colStdName).value))
                If stdName <> "" And Not stdNameDict.Exists(stdName) Then
                    stdNameDict.Add stdName, True
                End If
                Exit For
            End If
        Next j
    Next r
    
    ' --- Build dropdown list ---
    If stdNameDict.Count = 0 Then
        ClearDropdownValidation targetCell
        GoTo CleanExit
    End If
    
    stdNames = stdNameDict.Keys
    dropdownList = Join(stdNames, ",")
    
    ' --- Apply validation to target cell ---
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=dropdownList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in PopulateStandardNameDropdownForBatch: " & Err.Description, vbCritical

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    If tempFilePath <> "" And Not fso Is Nothing Then
        On Error Resume Next
        fso.DeleteFile tempFilePath, True
        On Error GoTo 0
    End If
    Set fso = Nothing
    Set qcDict = Nothing
    Set stdNameDict = Nothing
    Set wsQC = Nothing
End Sub

'-------------------------------------------------------------------------------
' PopulateTransactionDropdownForReversal
'-------------------------------------------------------------------------------
' Purpose: Populate Inventory Transaction dropdown filtered by Batch and RM Code
' Shows Trans_IDs where Trans_Type=CON and Reversed_Flag<>"Y"
'-------------------------------------------------------------------------------
Public Sub PopulateTransactionDropdownForReversal(ByVal batchNumber As String, ByVal rmCode As String, ByVal targetCell As Range)
    Dim conn As Object
    Dim rs As Object
    Dim fso As Object
    Dim tempFilePath As String
    Dim tempFolder As String
    Dim qcToRMDict As Object
    Dim transIDList As String
    Dim wsQC As Worksheet
    Dim colQCNumber As Long
    Dim colRMCode As Long
    Dim lastRow As Long
    Dim r As Long
    Dim qcNum As String
    Dim rmCodeVal As String
    Dim transIDVal As String
    Dim sqlQuery As String
    Dim qcList As String
    Dim qcKeys As Variant
    Dim j As Long
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set qcToRMDict = CreateObject("Scripting.Dictionary")
    
    ' --- Build QC Number to RM Code mapping from QC_Number worksheet ---
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colRMCode = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    
    If colQCNumber = 0 Or colRMCode = 0 Then
        MsgBox "Required columns not found in QC_Number worksheet.", vbCritical, "PopulateTransactionDropdownForReversal"
        GoTo CleanExit
    End If
    
    lastRow = wsQC.Cells(wsQC.Rows.Count, colQCNumber).End(xlUp).Row
    
    For r = 2 To lastRow
        qcNum = Trim$(CStr(wsQC.Cells(r, colQCNumber).value))
        rmCodeVal = Trim$(CStr(wsQC.Cells(r, colRMCode).value))
        
        If qcNum <> "" And LCase$(rmCodeVal) = LCase$(rmCode) Then
            If Not qcToRMDict.Exists(LCase$(qcNum)) Then
                qcToRMDict.Add LCase$(qcNum), True
            End If
        End If
    Next r
    
    If qcToRMDict.Count = 0 Then
        ClearDropdownValidation targetCell
        GoTo CleanExit
    End If
    
    tempFilePath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempFilePath = "" Then
        MsgBox "Unable to create temporary copy of RM_Trans.csv.", vbExclamation, "PopulateTransactionDropdownForReversal"
        GoTo CleanExit
    End If
    
    tempFolder = fso.GetParentFolderName(tempFilePath)
    
    Set conn = OpenADOConnection(tempFolder)
    
    If conn Is Nothing Then
        MsgBox "Unable to establish database connection.", vbCritical, "PopulateTransactionDropdownForReversal"
        GoTo CleanExit
    End If
    
    qcKeys = qcToRMDict.Keys
    qcList = ""
    For j = LBound(qcKeys) To UBound(qcKeys)
        If qcList <> "" Then qcList = qcList & ","
        qcList = qcList & "'" & Replace(CStr(qcKeys(j)), "'", "''") & "'"
    Next j
    
    sqlQuery = "SELECT " & RMT_HEADER_TRANS_ID & " " & _
               "FROM [" & fso.GetFileName(tempFilePath) & "] " & _
               "WHERE " & RMT_HEADER_FG_BATCH & " = '" & Replace(batchNumber, "'", "''") & "' " & _
               "AND " & RMT_HEADER_TRANS_TYPE & " = '" & TRANS_TYPE_CONSUMED & "' " & _
               "AND (" & RMT_HEADER_REVERSED_FLAG & " IS NULL OR " & RMT_HEADER_REVERSED_FLAG & " <> 'Y') " & _
               "AND LCASE(" & RMT_HEADER_QC_NUMBER & ") IN (" & qcList & ")"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn
    
    transIDList = ""
    If Not rs.EOF Then
        Do While Not rs.EOF
            transIDVal = Trim$(GetFieldValue(rs, RMT_HEADER_TRANS_ID) & "")
            
            If transIDVal <> "" Then
                If transIDList = "" Then
                    transIDList = transIDVal
                Else
                    transIDList = transIDList & "," & transIDVal
                End If
            End If
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    ' --- Apply validation ---
    If transIDList = "" Then
        ClearDropdownValidation targetCell
        GoTo CleanExit
    End If
    
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=transIDList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in PopulateTransactionDropdownForReversal: " & Err.Description, vbCritical

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    If tempFilePath <> "" And Not fso Is Nothing Then
        On Error Resume Next
        fso.DeleteFile tempFilePath, True
        On Error GoTo 0
    End If
    Set fso = Nothing
    Set qcToRMDict = Nothing
    Set wsQC = Nothing
End Sub

'-------------------------------------------------------------------------------
' ClearDropdownValidation
'-------------------------------------------------------------------------------
' Purpose: Remove data validation from a specified cell
'-------------------------------------------------------------------------------
Public Sub ClearDropdownValidation(ByVal targetCell As Range)
    On Error Resume Next
    targetCell.Validation.Delete
    targetCell.value = ""
    On Error GoTo 0
End Sub
