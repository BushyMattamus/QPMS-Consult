Attribute VB_Name = "BusLogReverseIssue"
' ========================================================================
' GOODS ISSUE REVERSAL BUSINESS LOGIC
' ========================================================================
' Handles reversal of goods issue (CON) transactions by creating COR
' transactions with negative quantities. Includes report display,
' validation, and updates to QC_Number balance and Nil_Stock flag.
' ========================================================================

Option Explicit

'-------------------------------------------------------------------------------
' SaveGoodsReversal
'-------------------------------------------------------------------------------
' Purpose: Post COR transaction to RM_Trans.csv and update related records
' Creates negative quantity transaction, marks original as reversed,
' updates QC balance, and handles Nil_Stock flag changes
'-------------------------------------------------------------------------------
Public Sub SaveGoodsReversal()
    Dim conn As Object
    Dim rs As Object
    Dim tempFilePath As String
    Dim originalTransID As String
    Dim reversalReason As String
    Dim qcNumber As String
    Dim qcStat As String
    Dim fgBatch As String
    Dim fgBatchStatus As String
    Dim quantity As Double
    Dim masterUoM As String
    Dim adjNote As String
    Dim qcTransID As String
    Dim piTransID As String
    Dim newBalance As Double
    Dim currentNilStock As String
    Dim reversedVal As String
    Dim userResponse As VbMsgBoxResult
    Dim sql As String
    Dim safeTransID As String
    
    On Error GoTo ErrorHandler
    
    ' --- 1. Get Transaction ID and Reason from form ---
    originalTransID = Trim$(GetFieldValByLabelC("Inventory Transaction:"))
    reversalReason = Trim$(GetFieldValByLabelC("Reason:"))
    
    ' Validate no commas
    If Not ValidateNoCommas(reversalReason, "Reason") Then GoTo CleanExit
    
    If originalTransID = "" Then
        MsgBox "Please select an Inventory Transaction to reverse.", vbExclamation, "SaveGoodsReversal"
        GoTo CleanExit
    End If
    
    If reversalReason = "" Then
        MsgBox "Please provide a reason for the reversal.", vbExclamation, "SaveGoodsReversal"
        GoTo CleanExit
    End If
    
    ' --- 2. Query original transaction from RM_Trans.csv using ADO ---
    tempFilePath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempFilePath = "" Then
        MsgBox "Unable to access RM_Trans.csv file.", vbCritical, "SaveGoodsReversal"
        GoTo CleanExit
    End If
    
    Set conn = OpenADOConnection(GetWorkbookPath() & "\Tables\temp\")
    
    If conn Is Nothing Then
        MsgBox "Unable to open database connection.", vbCritical, "SaveGoodsReversal"
        GoTo CleanExit
    End If
    
    safeTransID = Replace(originalTransID, "'", "''")
    
    sql = "SELECT " & _
          RMT_HEADER_TRANS_ID & ", " & _
          RMT_HEADER_QC_NUMBER & ", " & _
          RMT_HEADER_QC_STAT & ", " & _
          RMT_HEADER_FG_BATCH & ", " & _
          RMT_HEADER_FG_BATCH_STATUS & ", " & _
          RMT_HEADER_QUANTITY & ", " & _
          RMT_HEADER_MASTER_UOM & ", " & _
          RMT_HEADER_QC_TRANS_ID & ", " & _
          RMT_HEADER_PI_TRANS_ID & ", " & _
          RMT_HEADER_REVERSED_FLAG & " " & _
          "FROM [" & CreateObject("Scripting.FileSystemObject").GetFileName(tempFilePath) & "] " & _
          "WHERE " & RMT_HEADER_TRANS_ID & " = '" & safeTransID & "'"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1
    
    If rs.EOF Then
        MsgBox "Transaction ID not found: " & originalTransID, vbCritical, "SaveGoodsReversal"
        GoTo CleanExit
    End If
    
    reversedVal = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_REVERSED_FLAG) & ""))
    
    If UCase$(reversedVal) = "Y" Then
        MsgBox "This transaction has already been reversed.", vbExclamation, "SaveGoodsReversal"
        GoTo CleanExit
    End If
    
    qcNumber = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_QC_NUMBER) & ""))
    qcStat = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_QC_STAT) & ""))
    fgBatch = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_FG_BATCH) & ""))
    fgBatchStatus = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_FG_BATCH_STATUS) & ""))
    quantity = CDbl(GetFieldValue(rs, RMT_HEADER_QUANTITY))
    masterUoM = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_MASTER_UOM) & ""))
    qcTransID = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_QC_TRANS_ID) & ""))
    piTransID = Trim$(CStr(GetFieldValue(rs, RMT_HEADER_PI_TRANS_ID) & ""))
    
    ' --- 3. Confirm reversal with user ---
    userResponse = MsgBox("Are you sure you want to reverse this transaction?" & vbCrLf & vbCrLf & _
                          "Transaction ID: " & originalTransID & vbCrLf & _
                          "QC Number: " & qcNumber & vbCrLf & _
                          "FG Batch: " & fgBatch & vbCrLf & _
                          "Quantity: " & Abs(quantity) & " " & masterUoM & vbCrLf & _
                          "Reason: " & reversalReason, _
                          vbYesNo + vbQuestion, "Confirm Reversal")
    
    If userResponse <> vbYes Then
        GoTo CleanExit
    End If
    
    ' --- 4. Create COR transaction ---
    adjNote = originalTransID & REVERSAL_NOTE_SUFFIX & ". " & reversalReason
    
    WriteRMConsumptionTransaction _
        transType:=TRANS_TYPE_CONSUMED_REVERSAL, _
        qcNumber:=qcNumber, _
        qcStat:=qcStat, _
        fgBatch:=fgBatch, _
        fgBatchStatus:=fgBatchStatus, _
        quantity:=-quantity, _
        masterUoM:=masterUoM, _
        adjNote:=adjNote, _
        qcTransID:=qcTransID, _
        piTransID:=piTransID, _
        reversedFlag:=""
    
    ' --- 5. Update original transaction Reversed_Flag ---
    UpdateOriginalTransactionReversedFlag originalTransID
    
    ' --- 6. Calculate new balance and update QC_Number ---
    newBalance = GetQCBalanceFromRMTrans(qcNumber)
    
    ' --- 7. Check and update Nil_Stock flag if needed ---
    currentNilStock = GetCurrentNilStockFlag(qcNumber)
    
    If currentNilStock = "Y" And newBalance > 0 Then
        ' Update Nil_Stock flag to "N" and write audit record
        UpdateQCNumberNilStockFlag qcNumber, "N", newBalance, originalTransID
    Else
        ' Update balance only
        UpdateQCNumberBalance qcNumber, newBalance
    End If
    
    ' --- 8. Reload QC Number data ---
    BusLogDataLoads.LoadQCNumberData
    
    MsgBox "Goods issue reversal completed successfully." & vbCrLf & _
           "Original Transaction: " & originalTransID & vbCrLf & _
           "QC Number: " & qcNumber & vbCrLf & _
           "Quantity Restored: " & quantity & " " & masterUoM, vbInformation, "SaveGoodsReversal"
    
    ' --- 9. Refresh report display ---
    Dim batchNumber As String
    Dim rmCode As String
    
    batchNumber = Trim$(GetFieldValByLabelC("Batch Number:"))
    rmCode = Trim$(GetFieldValByLabelC("RM Code:"))
    
    If batchNumber <> "" And rmCode <> "" Then
        DisplayReversalTransactionReport batchNumber, rmCode
    End If
    
    ' --- 10. Clear form fields ---
    ClearReversalFormFields
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in SaveGoodsReversal: " & Err.Number & " - " & Err.Description, vbCritical, "SaveGoodsReversal"

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    If tempFilePath <> "" Then
        On Error Resume Next
        Kill tempFilePath
        On Error GoTo 0
    End If
End Sub

'-------------------------------------------------------------------------------
' DisplayReversalTransactionReport
'-------------------------------------------------------------------------------
' Purpose: Display report of CON transactions for batch/RM combination
' Writes to columns H:L starting at row 10
'-------------------------------------------------------------------------------
Public Sub DisplayReversalTransactionReport(ByVal batchNumber As String, ByVal rmCode As String)
    Dim transData As Variant
    Dim ws As Worksheet
    Dim startRow As Long
    Dim i As Long
    Dim lastReportRow As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    startRow = 10
    
    ' --- Clear existing report data (preserve headers in row 10) ---
    lastReportRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    If lastReportRow >= startRow Then
        ws.Range("H" & startRow & ":L" & lastReportRow).ClearContents
    End If
    
    ' --- Get transaction data ---
    transData = GetTransactionsForBatchAndRM(batchNumber, rmCode)
    
    If Not IsArray(transData) Then
        GoTo CleanExit
    End If
    
    If UBound(transData, 1) < 1 Then
        GoTo CleanExit
    End If
    
    ' --- Populate report ---
        
    For i = 1 To UBound(transData, 1)
        ws.Cells(startRow + i - 1, "H").value = transData(i, 1)
        ws.Cells(startRow + i - 1, "I").value = transData(i, 2)
        ws.Cells(startRow + i - 1, "J").value = transData(i, 3)
        ws.Cells(startRow + i - 1, "K").value = transData(i, 4)
        ws.Cells(startRow + i - 1, "L").value = transData(i, 5)
    Next i
        
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in DisplayReversalTransactionReport: " & Err.Number & " - " & Err.Description, vbCritical, "DisplayReversalTransactionReport"

CleanExit:
    Set ws = Nothing
End Sub

'-------------------------------------------------------------------------------
' UpdateOriginalTransactionReversedFlag
'-------------------------------------------------------------------------------
' Purpose: Mark original CON transaction as reversed in RM_Trans.csv
'-------------------------------------------------------------------------------
Private Sub UpdateOriginalTransactionReversedFlag(ByVal originalTransID As String)
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim lineParts() As String
    Dim headerParts() As String
    Dim newFileContent As String
    Dim i As Long
    Dim j As Long
    Dim idxTransID As Long
    Dim idxReversedFlag As Long
    Dim cleanHeader As String
    Dim transIDVal As String
    Dim updatedLine As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    ' --- Read file ---
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    Set ts = Nothing
    
    lines = Split(ToLF(fileContent), vbLf)
    
    ' --- Remove BOM if present (UTF-8 BOM: EF BB BF appears as ï»¿ when read as ANSI) ---
    If Len(lines(0)) >= 3 Then
        If Left$(lines(0), 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
            lines(0) = Mid$(lines(0), 4)
        End If
    End If
    
    ' --- Parse header ---
    headerParts = Split(lines(0), ",")
    idxTransID = -1
    idxReversedFlag = -1
    
    For j = LBound(headerParts) To UBound(headerParts)
        cleanHeader = Trim$(Replace(Replace(headerParts(j), """", ""), vbCr, ""))
        
        Select Case cleanHeader
            Case RMT_HEADER_TRANS_ID
                idxTransID = j
            Case RMT_HEADER_REVERSED_FLAG
                idxReversedFlag = j
        End Select
    Next j
    
    If idxTransID = -1 Or idxReversedFlag = -1 Then
        MsgBox "Required columns not found for flag update.", vbCritical, "UpdateOriginalTransactionReversedFlag"
        GoTo CleanExit
    End If
    
    ' --- Update matching line ---
    newFileContent = lines(0)
    
    For i = 1 To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            lineParts = Split(lines(i), ",")
            
            If UBound(lineParts) >= idxTransID Then
                transIDVal = Trim$(Replace(Replace(lineParts(idxTransID), """", ""), vbCr, ""))
                
                If transIDVal = originalTransID Then
                    lineParts(idxReversedFlag) = """Y"""
                    updatedLine = Join(lineParts, ",")
                    newFileContent = newFileContent & vbLf & updatedLine
                Else
                    newFileContent = newFileContent & vbLf & lines(i)
                End If
            Else
                newFileContent = newFileContent & vbLf & lines(i)
            End If
        End If
    Next i
    
    ' --- Write file ---
    Set ts = fso.CreateTextFile(filePath, True)
    ts.Write newFileContent & vbCrLf
    ts.Close
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in UpdateOriginalTransactionReversedFlag: " & Err.Number & " - " & Err.Description, vbCritical, "UpdateOriginalTransactionReversedFlag"

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

'-------------------------------------------------------------------------------
' GetCurrentNilStockFlag
'-------------------------------------------------------------------------------
' Purpose: Get current Nil_Stock flag value from QC_Numbers worksheet
'-------------------------------------------------------------------------------
Private Function GetCurrentNilStockFlag(ByVal qcNumber As String) As String
    Dim wsQC As Worksheet
    Dim colQCNumber As Long
    Dim colNilStock As Long
    Dim lastRow As Long
    Dim r As Long
    Dim qcVal As String
    
    On Error GoTo ErrorHandler
    
    Set wsQC = ThisWorkbook.Worksheets("QC_Numbers")
    
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colNilStock = FindHeaderCol(wsQC, 1, HEADER_NIL_STOCK_FLAG)
    
    If colQCNumber = 0 Or colNilStock = 0 Then
        GetCurrentNilStockFlag = ""
        GoTo CleanExit
    End If
    
    lastRow = wsQC.Cells(wsQC.Rows.Count, colQCNumber).End(xlUp).Row
    
    For r = 2 To lastRow
        qcVal = Trim$(CStr(wsQC.Cells(r, colQCNumber).value))
        
        If LCase$(qcVal) = LCase$(qcNumber) Then
            GetCurrentNilStockFlag = Trim$(CStr(wsQC.Cells(r, colNilStock).value))
            GoTo CleanExit
        End If
    Next r
    
    GetCurrentNilStockFlag = ""
    
    GoTo CleanExit

ErrorHandler:
    GetCurrentNilStockFlag = ""

CleanExit:
    Set wsQC = Nothing
End Function

'-------------------------------------------------------------------------------
' UpdateQCNumberNilStockFlag
'-------------------------------------------------------------------------------
' Purpose: Update Nil_Stock flag and balance in QC_Number.csv
' Also writes audit record to QC_Number_Changes.csv
'-------------------------------------------------------------------------------
Private Sub UpdateQCNumberNilStockFlag(ByVal qcNumber As String, ByVal newNilStock As String, ByVal newBalance As Double, ByVal refTransID As String)
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim lineParts() As String
    Dim headerParts() As String
    Dim line As String
    Dim i As Long
    Dim j As Long
    Dim idxQCNumber As Long
    Dim idxNilStock As Long
    Dim idxNilStockBy As Long
    Dim idxNilStockDate As Long
    Dim idxQCBalance As Long
    Dim idxChangeReason As Long
    Dim idxTransDate As Long
    Dim idxTransUser As Long
    Dim qcVal As String
    Dim found As Boolean
    Dim changeReason As String
    Dim cleanHeader As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not FileExists(filePath) Then
        MsgBox "QC_Number.csv file not found.", vbCritical, "UpdateQCNumberNilStockFlag"
        GoTo CleanExit
    End If
    
    ' --- Read file ---
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    Set ts = Nothing
    
    lines = Split(ToLF(fileContent), vbLf)
    
    ' --- Remove BOM if present ---
    If Len(lines(0)) >= 3 Then
        If Left$(lines(0), 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
            lines(0) = Mid$(lines(0), 4)
        End If
    End If
    
    ' --- Parse header from file content ---
    headerParts = Split(lines(0), ",")
    idxQCNumber = -1
    idxQCBalance = -1
    idxNilStock = -1
    idxNilStockBy = -1
    idxNilStockDate = -1
    idxChangeReason = -1
    idxTransDate = -1
    idxTransUser = -1
    
    For j = LBound(headerParts) To UBound(headerParts)
        cleanHeader = Trim$(Replace(Replace(headerParts(j), """", ""), vbCr, ""))
        
        Select Case cleanHeader
            Case HEADER_QC_NUMBER
                idxQCNumber = j
            Case HEADER_QC_BALANCE
                idxQCBalance = j
            Case HEADER_NIL_STOCK_FLAG
                idxNilStock = j
            Case HEADER_NIL_STOCK_FLAG_BY
                idxNilStockBy = j
            Case HEADER_NIL_STOCK_FLAG_DATE
                idxNilStockDate = j
            Case HEADER_CHANGE_REASON
                idxChangeReason = j
            Case HEADER_TRANS_DATE
                idxTransDate = j
            Case HEADER_TRANS_USER
                idxTransUser = j
        End Select
    Next j
    
    If idxQCNumber = -1 Or idxQCBalance = -1 Or idxNilStock = -1 Then
        MsgBox "Required columns not found in QC_Number.csv.", vbCritical, "UpdateQCNumberNilStockFlag"
        GoTo CleanExit
    End If
    
    ' --- Build change reason ---
    changeReason = "Nil Stock = " & newNilStock & ", Issue Reversal " & refTransID
    
    ' --- Find and update matching line ---
    found = False
    
    For i = 1 To UBound(lines)
        line = lines(i)
        If Trim$(line) <> "" Then
            lineParts = Split(line, ",")
            
            If UBound(lineParts) >= idxQCNumber Then
                qcVal = Trim$(Replace(lineParts(idxQCNumber), """", ""))
                
                If LCase$(qcVal) = LCase$(qcNumber) Then
                    lineParts(idxQCBalance) = """" & CStr(Round(newBalance, 3)) & """"
                    lineParts(idxNilStock) = """" & newNilStock & """"
                    
                    If idxNilStockBy > -1 And UBound(lineParts) >= idxNilStockBy Then
                        lineParts(idxNilStockBy) = """"""
                    End If
                    
                    If idxNilStockDate > -1 And UBound(lineParts) >= idxNilStockDate Then
                        lineParts(idxNilStockDate) = """"""
                    End If
                    
                    If idxChangeReason > -1 And UBound(lineParts) >= idxChangeReason Then
                        lineParts(idxChangeReason) = """" & changeReason & """"
                    End If
                    
                    If idxTransDate > -1 And UBound(lineParts) >= idxTransDate Then
                        lineParts(idxTransDate) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"
                    End If
                    
                    If idxTransUser > -1 And UBound(lineParts) >= idxTransUser Then
                        lineParts(idxTransUser) = """" & Environ$("Username") & """"
                    End If
                    
                    lines(i) = Join(lineParts, ",")
                    found = True
                    
                    ' --- Write audit record to QC_Number_Changes.csv ---
                    WriteQCNumberChangeAudit lines(i), refTransID
                    
                    Exit For
                End If
            End If
        End If
    Next i
    
    If Not found Then
        MsgBox "QC Number " & qcNumber & " not found in QC_Number.csv.", vbExclamation, "UpdateQCNumberNilStockFlag"
        GoTo CleanExit
    End If
    
    ' --- Write updated file ---
    Set ts = fso.OpenTextFile(filePath, 2, False)
    ts.Write Join(lines, vbCrLf)
    ts.Close
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in UpdateQCNumberNilStockFlag: " & Err.Number & " - " & Err.Description, vbCritical, "UpdateQCNumberNilStockFlag"

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

'-------------------------------------------------------------------------------
' UpdateQCNumberBalance
'-------------------------------------------------------------------------------
' Purpose: Update only the QC balance in QC_Number.csv (no Nil_Stock change)
'-------------------------------------------------------------------------------
Private Sub UpdateQCNumberBalance(ByVal qcNumber As String, ByVal newBalance As Double)
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim lineParts() As String
    Dim headerParts() As String
    Dim line As String
    Dim cleanHeader As String
    Dim i As Long
    Dim j As Long
    Dim idxQCNumber As Long
    Dim idxQCBalance As Long
    Dim idxNilStock As Long
    Dim idxNilStockBy As Long
    Dim idxNilStockDate As Long
    Dim idxChangeReason As Long
    Dim idxTransDate As Long
    Dim idxTransUser As Long
    Dim qcVal As String
    Dim found As Boolean
    Dim currentNilStock As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not FileExists(filePath) Then
        MsgBox "QC_Number.csv file not found.", vbCritical, "UpdateQCNumberBalance"
        GoTo CleanExit
    End If
    
    ' --- Read file ---
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    Set ts = Nothing
    
    lines = Split(ToLF(fileContent), vbLf)
    
    ' --- Remove BOM if present ---
    If Len(lines(0)) >= 3 Then
        If Left$(lines(0), 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
            lines(0) = Mid$(lines(0), 4)
        End If
    End If
    
    ' --- Parse header from file content ---
    headerParts = Split(lines(0), ",")
    idxQCNumber = -1
    idxQCBalance = -1
    idxNilStock = -1
    idxNilStockBy = -1
    idxNilStockDate = -1
    idxChangeReason = -1
    idxTransDate = -1
    idxTransUser = -1
    
    For j = LBound(headerParts) To UBound(headerParts)
        cleanHeader = Trim$(Replace(Replace(headerParts(j), """", ""), vbCr, ""))
        
        Select Case cleanHeader
            Case HEADER_QC_NUMBER
                idxQCNumber = j
            Case HEADER_QC_BALANCE
                idxQCBalance = j
            Case HEADER_NIL_STOCK_FLAG
                idxNilStock = j
            Case HEADER_NIL_STOCK_FLAG_BY
                idxNilStockBy = j
            Case HEADER_NIL_STOCK_FLAG_DATE
                idxNilStockDate = j
            Case HEADER_CHANGE_REASON
                idxChangeReason = j
            Case HEADER_TRANS_DATE
                idxTransDate = j
            Case HEADER_TRANS_USER
                idxTransUser = j
        End Select
    Next j
    
    If idxQCNumber = -1 Or idxQCBalance = -1 Then
        MsgBox "Required columns not found in QC_Number.csv.", vbCritical, "UpdateQCNumberBalance"
        GoTo CleanExit
    End If
    
    ' --- Find and update matching line ---
    found = False
    
    For i = 1 To UBound(lines)
        line = lines(i)
        If Trim$(line) <> "" Then
            lineParts = Split(line, ",")
            
            If UBound(lineParts) >= idxQCNumber Then
                qcVal = Trim$(Replace(lineParts(idxQCNumber), """", ""))
                
                If LCase$(qcVal) = LCase$(qcNumber) Then
                    lineParts(idxQCBalance) = """" & CStr(Round(newBalance, 3)) & """"
                    
                    ' --- Check if Nil_Stock was Y and balance is now positive ---
                    If idxNilStock > -1 And UBound(lineParts) >= idxNilStock Then
                        currentNilStock = Trim$(Replace(lineParts(idxNilStock), """", ""))
                        
                        If UCase$(currentNilStock) = "Y" And newBalance > 0 Then
                            lineParts(idxNilStock) = """N"""
                            
                            If idxNilStockBy > -1 And UBound(lineParts) >= idxNilStockBy Then
                                lineParts(idxNilStockBy) = """"""
                            End If
                            
                            If idxNilStockDate > -1 And UBound(lineParts) >= idxNilStockDate Then
                                lineParts(idxNilStockDate) = """"""
                            End If
                        End If
                    End If
                    
                    If idxChangeReason > -1 And UBound(lineParts) >= idxChangeReason Then
                        lineParts(idxChangeReason) = """Issue Reversal"""
                    End If
                    
                    If idxTransDate > -1 And UBound(lineParts) >= idxTransDate Then
                        lineParts(idxTransDate) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"
                    End If
                    
                    If idxTransUser > -1 And UBound(lineParts) >= idxTransUser Then
                        lineParts(idxTransUser) = """" & Environ$("Username") & """"
                    End If
                    
                    lines(i) = Join(lineParts, ",")
                    found = True
                    
                    ' --- Write audit record to QC_Number_Changes.csv ---
                    WriteQCNumberChangeAudit lines(i), "Balance Update"
                    
                    Exit For
                End If
            End If
        End If
    Next i
    
    If Not found Then
        MsgBox "QC Number '" & qcNumber & "' not found in QC_Number.csv.", vbExclamation, "UpdateQCNumberBalance"
        GoTo CleanExit
    End If
    
    ' --- Write updated file ---
    Set ts = fso.OpenTextFile(filePath, 2, False)
    ts.Write Join(lines, vbCrLf)
    ts.Close
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in UpdateQCNumberBalance: " & Err.Number & " - " & Err.Description, vbCritical, "UpdateQCNumberBalance"

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

'-------------------------------------------------------------------------------
' WriteQCNumberChangeAudit
'-------------------------------------------------------------------------------
' Purpose: Write audit record to QC_Number_Changes.csv
' Writes the full updated line from QC_Number.csv
'-------------------------------------------------------------------------------
Private Sub WriteQCNumberChangeAudit(ByVal updatedLine As String, ByVal refTransID As String)
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    
    Set ts = fso.OpenTextFile(filePath, 8, True)
    ts.WriteLine updatedLine
    ts.Close
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in WriteQCNumberChangeAudit: " & Err.Number & " - " & Err.Description, vbCritical, "WriteQCNumberChangeAudit"

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub
