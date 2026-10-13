Attribute VB_Name = "BusLogRecordLoader"
' ========================================================================
' RECORD LOADER FUNCTIONS
' ========================================================================
' Loads QC Number records, batch details, and RM information from
' worksheets and CSV files. Populates form fields and list controls.
' ========================================================================

Option Explicit

Public Sub LoadQCNumberRecord()
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim qcNumber As String
    Dim csvPath As String
    Dim fileNum As Integer
    Dim headerLine As String
    Dim dataLine As String
    Dim headers() As String
    Dim values() As String
    Dim foundRecord As Boolean
    Dim i As Integer
    Dim lastRetestDate As String
    Dim firstRetestDate As String
    Dim cleanValue As String
    Dim currentBalance As Double
    Dim retestDateValue As String
    
    ' --- Step 1: Get QC number from form ---
    qcNumber = GetFieldValByLabelC("QC Number:")
    
    If qcNumber = "" Then
        MsgBox "Error: Please enter a QC Number in cell D3", vbExclamation, "LoadQCNumberRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 2: Validate database file exists ---
    csvPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not FileExists(csvPath) Then
        MsgBox "Error: QC Number database file not found: " & csvPath, vbExclamation, "LoadQCNumberRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 3: Read CSV headers ---
    fileNum = FreeFile
    Open csvPath For Input As #fileNum
    Line Input #fileNum, headerLine
    headers = Split(headerLine, ",")
    
    ' --- Step 4: Search for matching QC number ---
    foundRecord = False
    Do Until EOF(fileNum)
        Line Input #fileNum, dataLine
        values = Split(dataLine, ",")
        
        If Len(values(0)) > 0 Then
            If Trim$(Replace(values(0), """", "")) = qcNumber Then
                foundRecord = True
                Exit Do
            End If
        End If
    Loop
    
    If Not foundRecord Then
        MsgBox "Error: QC Number " & qcNumber & " not found in database", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- Step 5: Clear destination cells ---
    Range("D4:D25").ClearContents
    
    ' --- Step 6: Map CSV values to form fields ---
    For i = 0 To UBound(headers)
        If i > UBound(values) Then Exit For
        
        cleanValue = Replace(values(i), """", "")
        
        Select Case Trim$(headers(i))
            Case HEADER_STANDARD_NAME
                Range("D4").value = cleanValue
            Case HEADER_RM_CODE
                Range("D5").value = Format$(cleanValue, "0000")
            Case HEADER_BRAND_NAME
                Range("D6").value = cleanValue
            Case HEADER_QUICK_CODE
                Range("D7").value = cleanValue
            Case HEADER_PART_ARTG
                Range("D8").value = cleanValue
            Case HEADER_VENDOR_NAME
                Range("D9").value = cleanValue
            Case HEADER_MANUFACTURER_NAME
                Range("D10").value = cleanValue
            Case HEADER_MANUFACTURER_LOT
                Range("D11").value = cleanValue
            Case HEADER_MANUFACTURER_EXPIRY
                SetDateStringInCell Range("D12"), ANSIToDate(cleanValue)
            Case HEADER_NUMBER_CONTAINERS
                Range("D13").value = cleanValue
            Case HEADER_MASTER_UOM
                Range("D15").value = cleanValue
            Case HEADER_QC_STATUS_FLAG
                Range("D16").value = cleanValue
            Case HEADER_LAST_RETEST_DATE
                lastRetestDate = ANSIToDate(cleanValue)
            Case HEADER_FIRST_RETEST_DATE
                firstRetestDate = ANSIToDate(cleanValue)
            Case HEADER_NEW_RETEST
                Range("D19").value = cleanValue
            Case HEADER_QC_TESTED
                Range("D17").value = cleanValue
            Case HEADER_TESTED_DATE
                SetDateStringInCell Range("D18"), ANSIToDate(cleanValue)
        End Select
    Next i
    
    ' --- Step 7: Get current balance from RM_Trans ---
    currentBalance = GetQCBalanceADO(qcNumber)
    Range("D14").value = currentBalance
    
    ' --- Step 8: Set retest date (last preferred over first) ---
    If Len(Trim$(lastRetestDate)) > 0 Then
        retestDateValue = lastRetestDate
    Else
        retestDateValue = firstRetestDate
    End If
    
    If Len(Trim$(retestDateValue)) > 0 Then
        SetDateStringInCell Range("D20"), retestDateValue
    Else
        Range("D20").value = ""
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LoadQCNumberRecord"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    Exit Sub
    
End Sub

Public Sub LoadBatchDetailUsage()
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim batchNumber As String
    Dim wsBatch As Worksheet
    Dim colBatchNumber As Long, colCPMU As Long, colDesc As Long, colExpiry As Long
    Dim lastRow As Long, r As Long
    Dim foundRecord As Boolean
    
    ' --- Step 1: Get batch number from form ---
    batchNumber = Trim$(Range("D3").value)
    
    If batchNumber = "" Then
        MsgBox "Error: Please enter a Batch Number in cell D3", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- Step 2: Reference Batch_Details worksheet ---
    Set wsBatch = ThisWorkbook.Worksheets("Batch_Details")
    
    ' --- Step 3: Find required column indices ---
    colBatchNumber = FindHeaderCol(wsBatch, 1, BX_HEADER_BATCH_NUMBER)
    colCPMU = FindHeaderCol(wsBatch, 1, BX_HEADER_PMI_NUMBER)
    colDesc = FindHeaderCol(wsBatch, 1, BX_HEADER_PRODUCT_DESCRIPTION)
    colExpiry = FindHeaderCol(wsBatch, 1, BX_HEADER_EXPIRY)
    
    If colBatchNumber = 0 Or colCPMU = 0 Or colDesc = 0 Or colExpiry = 0 Then
        MsgBox "Error: Required columns not found in Batch_Details sheet", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- Step 4: Search for batch number in worksheet ---
    lastRow = wsBatch.Cells(wsBatch.Rows.count, colBatchNumber).End(xlUp).row
    foundRecord = False
    
    For r = 2 To lastRow
        If Trim$(wsBatch.Cells(r, colBatchNumber).value) = batchNumber Then
            foundRecord = True
            
            ' --- Step 5: Clear and populate destination cells ---
            Range("D4:D6").ClearContents
            Range("D4").value = wsBatch.Cells(r, colCPMU).value
            Range("D5").value = wsBatch.Cells(r, colDesc).value
            Range("D6").value = wsBatch.Cells(r, colExpiry).value
            
            Exit For
        End If
    Next r
    
    If Not foundRecord Then
        MsgBox "Error: Batch Number " & batchNumber & " not found in Batch_Details", vbExclamation
        Range("D4:D6").ClearContents
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LoadBatchDetailUsage"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Set wsBatch = Nothing
    Exit Sub
    
End Sub

Public Sub PopulateQCNumberList(sourceCell As Range, Optional clearRange As Range)
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim rmCode As String
    Dim wsQC As Worksheet
    Dim wsIssuing As Worksheet
    Dim lastRow As Long, r As Long, destRow As Long
    Dim colQCNumber As Long, colRMCode As Long, colRetestDate As Long
    Dim colNilStockFlag As Long, colQCBalance As Long, colMasterUoM As Long, colQCStatus As Long
    Dim nilStockValue As String
    
    ' --- Step 1: Initialize worksheets ---
    Set wsIssuing = sourceCell.Parent
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    ' --- Step 2: Get RM Code from source cell ---
    rmCode = Trim$(CStr(sourceCell.value))
    
    If rmCode = "" Then
        GoTo CleanExit
    End If
    
    ' --- Step 3: Clear existing list and set destination row ---
    If Not clearRange Is Nothing Then
        clearRange.ClearContents
        destRow = clearRange.Rows(1).row
    Else
        wsIssuing.Range("H10:M100000").ClearContents
        destRow = 10
    End If
    
    ' --- Step 4: Find required columns in QC_Number sheet ---
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colRMCode = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    colRetestDate = FindHeaderCol(wsQC, 1, HEADER_LAST_RETEST_DATE)
    colNilStockFlag = FindHeaderCol(wsQC, 1, HEADER_NIL_STOCK_FLAG)
    colQCBalance = FindHeaderCol(wsQC, 1, HEADER_QC_BALANCE)
    colMasterUoM = FindHeaderCol(wsQC, 1, HEADER_MASTER_UOM)
    colQCStatus = FindHeaderCol(wsQC, 1, HEADER_QC_STATUS_FLAG)
    
    If colQCNumber = 0 Or colRMCode = 0 Or colRetestDate = 0 Or _
       colNilStockFlag = 0 Or colQCBalance = 0 Or colMasterUoM = 0 Or colQCStatus = 0 Then
        MsgBox "Error: Required columns not found in QC_Number sheet", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- Step 5: Search and populate matching RM Code records ---
    lastRow = wsQC.Cells(wsQC.Rows.count, colRMCode).End(xlUp).row
    
    For r = 2 To lastRow
        If Trim$(wsQC.Cells(r, colRMCode).value) = rmCode Then
            wsIssuing.Cells(destRow, "H").value = wsQC.Cells(r, colQCNumber).value
            wsIssuing.Cells(destRow, "I").value = wsQC.Cells(r, colRetestDate).value
            
            nilStockValue = Trim$(wsQC.Cells(r, colNilStockFlag).value)
            If nilStockValue = "" Then nilStockValue = "N"
            wsIssuing.Cells(destRow, "J").value = ToYesNo(nilStockValue)
            
            wsIssuing.Cells(destRow, "K").value = wsQC.Cells(r, colQCBalance).value
            wsIssuing.Cells(destRow, "L").value = wsQC.Cells(r, colMasterUoM).value
            wsIssuing.Cells(destRow, "M").value = wsQC.Cells(r, colQCStatus).value
            
            destRow = destRow + 1
        End If
    Next r
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PopulateQCNumberList"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Set wsQC = Nothing
    Set wsIssuing = Nothing
    Exit Sub

End Sub

Public Sub LoadRMDetailsFromQC(sourceCell As Range, descOutputCell As Range, codeOutputCell As Range)
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim qcNumber As String
    Dim wsQC As Worksheet
    Dim colQCNumber As Long, colRMCode As Long, colStandardName As Long
    Dim lastRow As Long, r As Long
    Dim foundRecord As Boolean
    
    ' --- Step 1: Get QC Number from source cell ---
    qcNumber = Trim$(CStr(sourceCell.value))
    
    If qcNumber = "" Then
        descOutputCell.ClearContents
        codeOutputCell.ClearContents
        GoTo CleanExit
    End If
    
    ' --- Step 2: Reference QC_Number worksheet ---
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    ' --- Step 3: Find required column indices ---
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colRMCode = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    colStandardName = FindHeaderCol(wsQC, 1, HEADER_STANDARD_NAME)
    
    If colQCNumber = 0 Or colRMCode = 0 Or colStandardName = 0 Then
        MsgBox "Error: Required columns not found in QC_Number sheet", vbExclamation
        descOutputCell.ClearContents
        codeOutputCell.ClearContents
        GoTo CleanExit
    End If
    
    ' --- Step 4: Search for QC Number in worksheet ---
    lastRow = wsQC.Cells(wsQC.Rows.count, colQCNumber).End(xlUp).row
    foundRecord = False
    
    For r = 2 To lastRow
        If Trim$(wsQC.Cells(r, colQCNumber).value) = qcNumber Then
            foundRecord = True
            
            ' --- Step 5: Populate output cells ---
            descOutputCell.value = wsQC.Cells(r, colStandardName).value
            codeOutputCell.value = Format$(wsQC.Cells(r, colRMCode).value, "0000")
            
            Exit For
        End If
    Next r
    
    If Not foundRecord Then
        MsgBox "Error: QC Number " & qcNumber & " not found in database", vbExclamation
        descOutputCell.ClearContents
        codeOutputCell.ClearContents
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LoadRMDetailsFromQC"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Set wsQC = Nothing
    Exit Sub
    
End Sub

Public Sub ResolveRMCodeFromDescription(sourceCell As Range, outputCell As Range)
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim desc As String
    Dim wsQC As Worksheet
    Dim colStd As Long, colRM As Long
    Dim lastRow As Long, r As Long
    Dim dict As Object
    Dim valStd As String, valRM As String
    
    ' --- Step 1: Get description from source cell ---
    desc = Trim$(CStr(sourceCell.value))
    
    If desc = "" Then
        GoTo CleanExit
    End If
    
    ' --- Step 2: Reference QC_Number worksheet ---
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    ' --- Step 3: Find required columns ---
    colStd = FindHeaderCol(wsQC, 1, HEADER_STANDARD_NAME)
    colRM = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    
    If colStd = 0 Or colRM = 0 Then
        GoTo CleanExit
    End If
    
    ' --- Step 4: Build dictionary of matching RM codes ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    lastRow = wsQC.Cells(wsQC.Rows.count, colStd).End(xlUp).row
    
    For r = 2 To lastRow
        valStd = Trim$(CStr(wsQC.Cells(r, colStd).value))
        If StrComp(valStd, desc, vbTextCompare) = 0 Then
            valRM = Trim$(CStr(wsQC.Cells(r, colRM).value))
            If valRM <> "" Then
                If Not dict.Exists(valRM) Then
                    dict.Add valRM, 1
                End If
            End If
        End If
    Next r
    
    ' --- Step 5: Output single match or prompt for selection ---
    If dict.count = 1 Then
        outputCell.NumberFormat = "@"
        outputCell.value = RM_Pad4(dict.keys()(0))
    Else
        outputCell.value = "Select from list"
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ResolveRMCodeFromDescription"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not dict Is Nothing Then Set dict = Nothing
    Set wsQC = Nothing
    Exit Sub
    
End Sub
