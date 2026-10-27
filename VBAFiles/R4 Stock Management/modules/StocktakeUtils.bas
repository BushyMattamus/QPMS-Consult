Attribute VB_Name = "StocktakeUtils"
' ========================================================================
' STOCKTAKE TRANSACTION UTILITIES
' ========================================================================
' Handles stocktake entry, QC counting, and physical inventory reconciliation
' ========================================================================

Option Explicit

' Returns True if given QC number already has a physical inventory entry in the same month as today
Public Function IsQCCountedThisMonth(ByVal qcNumber As String) As Boolean
    Dim fso As Object, ts As Object
    Dim filePath As String, fileContent As String
    Dim lines() As String, i As Long
    Dim header As String, parts() As String
    Dim idxCode As Long, idxDate As Long
    Dim hParts() As String
    Dim dt As Date
    Dim cellVal As String
    
    On Error GoTo ErrHandler
    IsQCCountedThisMonth = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_PHYSICAL_INVENTORY)
    If Not fso.FileExists(filePath) Then
        GoTo CleanExit
    End If
    
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    If Trim$(fileContent) = "" Then
        GoTo CleanExit
    End If
    
    lines = Split(fileContent, vbCrLf)
    hParts = Split(lines(0), ",")
    idxCode = -1: idxDate = -1
    Dim j As Long
    For j = 0 To UBound(hParts)
        Dim h As String: h = Replace(hParts(j), """", "")
        If LCase$(h) = "code" Or LCase$(h) = "qc_number" Then idxCode = j
        If LCase$(h) = "counted_date" Or LCase$(h) = "date" Then idxDate = j
    Next j
    If idxCode = -1 Or idxDate = -1 Then
        GoTo CleanExit
    End If
    
    For i = 1 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            parts = Split(lines(i), ",")
            If UBound(parts) >= idxCode Then
                cellVal = Replace(parts(idxCode), """", "")
                If StrComp(Trim$(cellVal), Trim$(qcNumber), vbTextCompare) = 0 Then
                    If UBound(parts) >= idxDate Then
                        On Error Resume Next
                        dt = CDate(Replace(parts(idxDate), """", ""))
                        On Error GoTo ErrHandler
                        If Year(dt) = Year(Date) And Month(dt) = Month(Date) Then
                            IsQCCountedThisMonth = True
                            GoTo CleanExit
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    GoTo CleanExit

ErrHandler:
    IsQCCountedThisMonth = False

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' Append a single stocktake entry and post immediate transaction (PAYG)
' Returns True on success
' ========================================================================
Public Function StocktakeQCNumEntry( _
    ByVal qcNumber As String, _
    ByVal countedQty As Double, _
    ByVal countedUoM As String, _
    ByVal countedBy As String, _
    Optional ByVal piNumber As String = "", _
    Optional ByVal newRetestDate As String = "") As Boolean

    Dim piNo As String
    Dim bookQty As Double
    Dim result As StockUpdateResult
    Dim reason As String
    Dim fso As Object
    Dim ts As Object
    Dim physPath As String
    Dim line As String
    Dim desc As String
    Dim rmCode As String
    Dim masterUoM As String
    Dim wtPerML As Double
    Dim convertedQty As Double
    Dim zeroAdjustment As Boolean
    Dim retestQCTransID As String
    Dim finalQCTransID As String
    Dim currentRetestDate As String
    Dim nilStockFlagDate As String
    Dim updateNilStock As Boolean
    Dim qcTransID As String
    Dim updateResult As Boolean
    Dim tBlock As Double
    
    On Error GoTo ErrHandler
    StocktakeQCNumEntry = False
    tBlock = PerfTimerNow()
    
    If IsQCCountedThisMonth(qcNumber) Then
        GoTo CleanExit
    End If
    
    PerfTimerLog "STOCKTAKE", "6a IsQCCountedThisMonth", tBlock
    tBlock = PerfTimerNow()

    If Trim$(piNumber) = "" Then
        piNo = GetNextPINumber()
    Else
        piNo = piNumber
    End If

    PerfTimerLog "STOCKTAKE", "6b GetNextPINumber (if no PI passed)", tBlock
    tBlock = PerfTimerNow()

    bookQty = GetQCBalanceADO(qcNumber)
    PerfTimerLog "STOCKTAKE", "6c Book balance (GetQCBalanceADO)", tBlock
    tBlock = PerfTimerNow()
           
    convertedQty = ConvertToMasterUoM(countedQty, countedUoM, masterUoM, wtPerML)
    PerfTimerLog "STOCKTAKE", "6d ConvertToMasterUoM", tBlock
    tBlock = PerfTimerNow()
    
    zeroAdjustment = (Round(convertedQty, 6) = Round(bookQty, 6))
    
    If Not zeroAdjustment Then
        reason = "Stocktake"
        
        result = ProcessStockUpdate( _
            qcNumber:=qcNumber, _
            inputQty:=countedQty, _
            inputUoM:=countedUoM, _
            reason:=reason, _
            adjustmentType:="", _
            confirmNilStock:=True, _
            fgBatch:="", _
            fgBatchStatus:="", _
            allowZeroQty:=True, _
            piNumber:=piNo _
        )
        
        If Not result.Success Then
            GoTo CleanExit
        End If
    Else
        nilStockFlagDate = GetQCNilStockFlagDate(qcNumber)
        updateNilStock = False
        
        If Round(convertedQty, 6) = 0 And nilStockFlagDate = "" Then
            updateNilStock = True
            qcTransID = GetNextQCTransactionID()
            
            updateResult = UpdateQCNumberRecord( _
                qcNumber:=qcNumber, _
                qcBalance:=0, _
                nilStockFlag:="Y", _
                nilStockFlagDate:=Format$(Date, "dd/mm/yyyy"), _
                nilStockFlagBy:=Environ$("Username"), _
                changeReason:=piNo & " Stocktake - Nil Stock (zero balance confirmed)", _
                qcTransID:=qcTransID)
            
            If Not updateResult Then
                MsgBox "Failed to update nil stock flag for QC Number " & qcNumber, vbCritical, "StocktakeQCNumEntry"
                GoTo CleanExit
            End If
        End If
    End If
    
    PerfTimerLog "STOCKTAKE", "6e Adjustment or nil stock update", tBlock
    tBlock = PerfTimerNow()

    If Trim$(newRetestDate) <> "" Then
        currentRetestDate = GetQCRetestDate(qcNumber)
        If DateToANSI(newRetestDate) <> currentRetestDate Then
            retestQCTransID = UpdateQCRetestAndExpiryDates(qcNumber, newRetestDate, piNo)
        End If
    End If
    
    PerfTimerLog "STOCKTAKE", "6f Retest/expiry update", tBlock
    tBlock = PerfTimerNow()

    If retestQCTransID <> "" Then
        finalQCTransID = retestQCTransID
    ElseIf updateNilStock And qcTransID <> "" Then
        finalQCTransID = qcTransID
    ElseIf Not zeroAdjustment And result.qcTransID <> "" Then
        finalQCTransID = result.qcTransID
    Else
        finalQCTransID = "NA"
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    physPath = GetWorkbookTableFilePath(FILE_PHYSICAL_INVENTORY)
    
    If Not fso.FileExists(physPath) Then
        Set ts = fso.OpenTextFile(physPath, 2, True)
        ts.WriteLine "Doc_Number,Counted_By,Counted_Date,Code,Description,Book_Qty,Counted_Qty,Diff_Qty,Trans_ID"
        ts.Close
        Set ts = Nothing
    End If

    desc = GetQCDescription(qcNumber)
    PerfTimerLog "STOCKTAKE", "6g GetQCDescription", tBlock
    tBlock = PerfTimerNow()
    
    line = """" & piNo & """," & _
           """" & countedBy & """," & _
           """" & Format$(Now, "yyyy-mm-dd HH:mm:ss") & """," & _
           """" & qcNumber & """," & _
           """" & Replace(desc, """", """""") & """," & _
           """" & Format$(bookQty, "0.000") & """," & _
           """" & Format$(convertedQty, "0.000") & """," & _
           """" & Format$(convertedQty - bookQty, "0.000") & """," & _
           """" & finalQCTransID & """"
    
    Set ts = fso.OpenTextFile(physPath, 8, True)
    ts.WriteLine line
    ts.Close
    
    PerfTimerLog "STOCKTAKE", "6h Physical inventory append", tBlock
    StocktakeQCNumEntry = True
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "StocktakeQCNumEntry"

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' Update QC_Number retest and expiry dates based on testing flag
' If Testing_Flag = "N": update Mfg_Expiry, First_Retest_Date, Last_Retest_Date
' If Testing_Flag = "Y": update only Last_Retest_Date
' ========================================================================
Public Function UpdateQCRetestAndExpiryDates(ByVal qcNumber As String, ByVal newRetestDate As String, ByVal piNumber As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim qcFilePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim parts() As String
    Dim i As Long
    Dim colQC As Long
    Dim colTesting As Long
    Dim testingFlag As String
    Dim ansiDate As String
    Dim found As Boolean
    Dim changeReason As String
    Dim qcNum As String
    Dim updateResult As Boolean
    Dim qcTransID As String
    
    On Error GoTo ErrHandler
    
    UpdateQCRetestAndExpiryDates = ""
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    qcFilePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not fso.FileExists(qcFilePath) Then
        MsgBox "QC_Number.csv file not found at " & qcFilePath, vbExclamation, "UpdateQCRetestAndExpiryDates"
        GoTo CleanExit
    End If
    
    Set ts = fso.OpenTextFile(qcFilePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    Set ts = Nothing
    
    If Trim$(fileContent) = "" Then
        MsgBox "QC_Number.csv file is empty.", vbExclamation, "UpdateQCRetestAndExpiryDates"
        GoTo CleanExit
    End If
    
    lines = Split(fileContent, vbCrLf)
    
    colQC = GetHeaderIndex0(HEADER_QC_NUMBER)
    colTesting = GetHeaderIndex0(HEADER_TESTING_FLAG)
    
    If colQC = -1 Or colTesting = -1 Then
        MsgBox "Required columns not found in QC_Number.csv.", vbExclamation, "UpdateQCRetestAndExpiryDates"
        GoTo CleanExit
    End If
    
    ansiDate = DateToANSI(newRetestDate)
    found = False
    
    For i = 1 To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            parts = ParseCSVLine(lines(i))
            
            If UBound(parts) >= colQC Then
                qcNum = Trim$(Replace(parts(colQC), """", ""))
                
                If StrComp(qcNum, qcNumber, vbTextCompare) = 0 Then
                    found = True
                    
                    If UBound(parts) >= colTesting Then
                        testingFlag = Trim$(UCase$(Replace(parts(colTesting), """", "")))
                        changeReason = piNumber & " Stocktake - Change Retest/Expiry "
                        qcTransID = GetNextQCTransactionID()
                        
                        If testingFlag = "N" Then
                            updateResult = UpdateQCNumberRecord( _
                                qcNumber:=qcNumber, _
                                manufacturerExpiry:=ansiDate, _
                                firstRetestDate:=ansiDate, _
                                lastRetestDate:=ansiDate, _
                                changeReason:=changeReason, _
                                qcTransID:=qcTransID _
                            )
                        Else
                            updateResult = UpdateQCNumberRecord( _
                                qcNumber:=qcNumber, _
                                lastRetestDate:=ansiDate, _
                                changeReason:=changeReason, _
                                qcTransID:=qcTransID _
                            )
                        End If
                        
                        If Not updateResult Then
                            MsgBox "Failed to update QC Number record.", vbCritical, "UpdateQCRetestAndExpiryDates"
                            qcTransID = ""
                        End If
                        
                        UpdateQCRetestAndExpiryDates = qcTransID
                    End If
                    
                    Exit For
                End If
            End If
        End If
    Next i
    
    If Not found Then
        MsgBox "QC Number " & qcNumber & " not found in QC_Number.csv.", vbExclamation, "UpdateQCRetestAndExpiryDates"
        GoTo CleanExit
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "UpdateQCRetestAndExpiryDates"
    UpdateQCRetestAndExpiryDates = ""
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' Batch entry: post multiple stocktake records as you go
' ========================================================================
Public Function StocktakeBatchEntry(records As Variant) As Object
    Dim report As Object
    Dim i As Long
    Dim rec As Variant
    Dim ok As Boolean
    
    Set report = CreateObject("Scripting.Dictionary")
    report("Total") = 0
    report("Posted") = 0
    report("Skipped_Duplicate") = 0
    report("Failed") = 0
    report("Errors") = CreateObject("Scripting.Dictionary")
    
    For i = LBound(records) To UBound(records)
        rec = records(i)
        report("Total") = report("Total") + 1
        On Error Resume Next
        ok = StocktakeQCNumEntry( _
                CStr(rec(0)), _
                CDbl(rec(1)), _
                CStr(rec(2)), _
                CStr(rec(3)) _
             )
        If Err.Number <> 0 Then
            report("Failed") = report("Failed") + 1
            report("Errors")(report("Total")) = "Error at record " & i & ": " & Err.Description
            Err.Clear
        Else
            If ok Then
                report("Posted") = report("Posted") + 1
            Else
                If IsQCCountedThisMonth(CStr(rec(0))) Then
                    report("Skipped_Duplicate") = report("Skipped_Duplicate") + 1
                Else
                    report("Failed") = report("Failed") + 1
                End If
            End If
        End If
        On Error GoTo 0
    Next i
    
    Set StocktakeBatchEntry = report
End Function

' ========================================================================
' Get last count date for each QC number from Physical_Inventory.csv
' Returns dictionary mapping QC_Number -> last count date (ANSI format)
' ========================================================================
Public Function GetLastCountDateLookup() As Object
    Dim dict As Object
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim i As Long
    Dim hParts() As String
    Dim parts() As String
    Dim idxCode As Long
    Dim idxDate As Long
    Dim qcNumber As String
    Dim countDate As String
    Dim ansiDate As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_PHYSICAL_INVENTORY)
    
    If Not fso.FileExists(filePath) Then
        GoTo CleanExit
    End If
    
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    Set ts = Nothing
    
    If Trim$(fileContent) = "" Then
        GoTo CleanExit
    End If
    
    lines = Split(fileContent, vbCrLf)
    
    If UBound(lines) < 1 Then
        GoTo CleanExit
    End If
    
    hParts = Split(lines(0), ",")
    idxCode = -1
    idxDate = -1
    
    Dim j As Long
    For j = 0 To UBound(hParts)
        Dim h As String
        h = LCase$(Replace(hParts(j), """", ""))
        If h = "code" Or h = "qc_number" Then idxCode = j
        If h = "counted_date" Or h = "date" Then idxDate = j
    Next j
    
    If idxCode = -1 Or idxDate = -1 Then
        GoTo CleanExit
    End If
    
    For i = 1 To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            parts = Split(lines(i), ",")
            
            If UBound(parts) >= idxCode And UBound(parts) >= idxDate Then
                qcNumber = Trim$(Replace(parts(idxCode), """", ""))
                countDate = Trim$(Replace(parts(idxDate), """", ""))
                
                If qcNumber <> "" And countDate <> "" Then
                    ansiDate = DateToANSI(countDate)
                    
                    If Not dict.Exists(qcNumber) Or ansiDate > CStr(dict(qcNumber)) Then
                        dict(qcNumber) = ansiDate
                    End If
                End If
            End If
        End If
    Next i
    
    GoTo CleanExit

ErrHandler:

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set GetLastCountDateLookup = dict
    Exit Function
End Function


' ========================================================================
' Helper: Get Nil Stock Flag Date from QC_Number worksheet
' ========================================================================
Private Function GetQCNilStockFlagDate(ByVal qcNumber As String) As String
    Dim ws As Worksheet
    Dim colQC As Long
    Dim colNilStockDate As Long
    Dim lastRow As Long
    Dim r As Long
    
    On Error GoTo ErrHandler
    GetQCNilStockFlagDate = ""
    
    Set ws = ThisWorkbook.Worksheets("QC_Number")
    colQC = FindHeaderCol(ws, 1, HEADER_QC_NUMBER)
    colNilStockDate = FindHeaderCol(ws, 1, HEADER_NIL_STOCK_FLAG_DATE)
    
    If colQC = 0 Or colNilStockDate = 0 Then
        GoTo CleanExit
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colQC).End(xlUp).row
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, colQC).value)) = qcNumber Then
            GetQCNilStockFlagDate = Trim$(CStr(ws.Cells(r, colNilStockDate).value))
            GoTo CleanExit
        End If
    Next r
    
    GoTo CleanExit

ErrHandler:
    GetQCNilStockFlagDate = ""

CleanExit:
    Set ws = Nothing
    Exit Function
End Function

' ========================================================================
' Returns a dictionary mapping QC_Number -> Boolean (True if counted this month)
' ========================================================================
Public Function GetStocktakeMonthStatus() As Object
    Dim dict As Object
    Dim fso As Object, ts As Object
    Dim filePath As String, fileContent As String
    Dim lines() As String, i As Long
    Dim hParts() As String, parts() As String
    Dim idxCode As Long, idxDate As Long
    Dim dt As Date
    
    Set dict = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_PHYSICAL_INVENTORY)
    If Not fso.FileExists(filePath) Then
        GoTo CleanExit
    End If
    
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    If Trim$(fileContent) = "" Then
        GoTo CleanExit
    End If
    
    lines = Split(fileContent, vbCrLf)
    hParts = Split(lines(0), ",")
    idxCode = -1: idxDate = -1
    For i = 0 To UBound(hParts)
        Dim h As String: h = LCase$(Replace(hParts(i), """", ""))
        If h = "code" Or h = "qc_number" Then idxCode = i
        If h = "counted_date" Or h = "date" Then idxDate = i
    Next i
    If idxCode = -1 Or idxDate = -1 Then
        GoTo CleanExit
    End If
    
    For i = 1 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            parts = Split(lines(i), ",")
            If UBound(parts) >= idxCode Then
                Dim qc As String: qc = Replace(parts(idxCode), """", "")
                On Error Resume Next
                dt = CDate(Replace(parts(idxDate), """", ""))
                On Error GoTo 0
                If Year(dt) = Year(Date) And Month(dt) = Month(Date) Then
                    If Not dict.Exists(qc) Then dict(qc) = True
                End If
            End If
        End If
    Next i
    
    GoTo CleanExit

ErrHandler:

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set GetStocktakeMonthStatus = dict
    Exit Function
End Function

' ========================================================================
' Helper: Get QC Description from QC_Number worksheet
' ========================================================================
Private Function GetQCDescription(ByVal qcNumber As String) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim colQC As Long, colDesc As Long, lastRow As Long, r As Long
    GetQCDescription = ""
    
    Set ws = ThisWorkbook.Worksheets("QC_Number")
    colQC = FindHeaderCol(ws, 1, HEADER_QC_NUMBER)
    colDesc = FindHeaderCol(ws, 1, "Description")
    
    If colQC = 0 Or colDesc = 0 Then
        GoTo CleanExit
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colQC).End(xlUp).row
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, colQC).value)) = qcNumber Then
            GetQCDescription = Trim$(CStr(ws.Cells(r, colDesc).value))
            GoTo CleanExit
        End If
    Next r
    
    GoTo CleanExit

ErrHandler:
    GetQCDescription = ""

CleanExit:
    Set ws = Nothing
    Exit Function
End Function

' ========================================================================
' Helper: Get current retest date from QC_Number record
' ========================================================================
Private Function GetQCRetestDate(ByVal qcNumber As String) As String
    Dim ws As Worksheet
    Dim colQC As Long
    Dim colRetest As Long
    Dim lastRow As Long
    Dim r As Long
    
    On Error GoTo ErrHandler
    GetQCRetestDate = ""
    
    Set ws = ThisWorkbook.Worksheets("QC_Number")
    colQC = FindHeaderCol(ws, 1, HEADER_QC_NUMBER)
    colRetest = FindHeaderCol(ws, 1, HEADER_LAST_RETEST_DATE)
    
    If colQC = 0 Or colRetest = 0 Then
        GoTo CleanExit
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colQC).End(xlUp).row
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, colQC).value)) = qcNumber Then
            GetQCRetestDate = Trim$(CStr(ws.Cells(r, colRetest).value))
            GoTo CleanExit
        End If
    Next r
    
    GoTo CleanExit

ErrHandler:
    GetQCRetestDate = ""

CleanExit:
    Set ws = Nothing
    Exit Function
End Function
