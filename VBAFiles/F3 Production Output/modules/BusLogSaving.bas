Attribute VB_Name = "BusLogSaving"
Option Explicit

Sub Save_ProdOutput()
    On Error GoTo ErrHandler
    
    Dim wsInput As Worksheet
    Dim rngProdLookup As Range
    Dim outputData(1 To 39) As Variant
    Dim quickCode As String
    Dim batchNumber As String
    Dim startRoom As String
    Dim finishRoom As String
    Dim plannedStartDate As Variant
    Dim actualStartDate As Variant
    Dim actualFinishDate As Variant
    Dim plannedReleaseQty As Variant
    Dim sampleQty As Variant
    Dim totalQty As Long
    Dim notesText As String
    Dim prepStaff1 As String
    Dim prepStaff2 As String
    Dim prepStaff3 As String
    Dim makeStaff1 As String
    Dim makeStaff2 As String
    Dim makeStaff3 As String
    Dim checkStaff1 As String
    Dim checkStaff2 As String
    Dim checkStaff3 As String
    Dim fillStaff1 As String
    Dim fillStaff2 As String
    Dim fillStaff3 As String
    Dim labelStaff1 As String
    Dim labelStaff2 As String
    Dim labelStaff3 As String
    Dim packStaff1 As String
    Dim packStaff2 As String
    Dim packStaff3 As String
    Dim otherStaff1 As String
    Dim otherStaff2 As String
    Dim otherStaff3 As String
    Dim productCategory As String
    Dim productName As String
    Dim productType As String
    Dim pmiCode As String
    Dim pmiVersion As String
    Dim response As VbMsgBoxResult
    Dim requiresStatusUpdate As Boolean
    Dim batchValidationResult As BatchStatusResult
    Dim prodMasterWs As Worksheet
    Dim prodMasterRange As Range
    Dim lookupRow As Range
    Dim colA As Long
    Dim colB As Long
    Dim colC As Long
    Dim colJ As Long
    Dim batchTransID As String
    Dim batchUpdateSuccess As Boolean
    
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    Set wsInput = ThisWorkbook.Sheets("Production Entry")
    
    quickCode = GetFieldValByLabelB("Quick Code:")
    productName = GetFieldValByLabelB("Product:")
    pmiCode = GetFieldValByLabelB("PMI Number:")
    pmiVersion = GetFieldValByLabelB("PMI Version:")
    batchNumber = GetFieldValByLabelB("Batch Number:")
    startRoom = GetFieldValByLabelB("Start Room:")
    finishRoom = GetFieldValByLabelB("Finish Room:")
    plannedStartDate = GetFieldValByLabelB("Planned Start Date:")
    actualStartDate = GetFieldValByLabelB("Actual Start Date:")
    actualFinishDate = GetFieldValByLabelB("Actual Finish Date:")
    plannedReleaseQty = GetFieldValByLabelB("Planned Release Qty:")
    sampleQty = GetFieldValByLabelB("Sample Qty:")
    notesText = GetFieldValByLabelB("Notes:")
    prepStaff1 = GetFieldValByLabelB("Production Preparation:")
    prepStaff2 = Trim(wsInput.Range("D23").value)
    prepStaff3 = Trim(wsInput.Range("E23").value)
    makeStaff1 = GetFieldValByLabelB("Make:")
    makeStaff2 = Trim(wsInput.Range("D24").value)
    makeStaff3 = Trim(wsInput.Range("E24").value)
    checkStaff1 = GetFieldValByLabelB("Checking:")
    checkStaff2 = Trim(wsInput.Range("D25").value)
    checkStaff3 = Trim(wsInput.Range("E25").value)
    fillStaff1 = GetFieldValByLabelB("Fill:")
    fillStaff2 = Trim(wsInput.Range("D26").value)
    fillStaff3 = Trim(wsInput.Range("E26").value)
    labelStaff1 = GetFieldValByLabelB("Label:")
    labelStaff2 = Trim(wsInput.Range("D27").value)
    labelStaff3 = Trim(wsInput.Range("E27").value)
    packStaff1 = GetFieldValByLabelB("Pack:")
    packStaff2 = Trim(wsInput.Range("D28").value)
    packStaff3 = Trim(wsInput.Range("E28").value)
    otherStaff1 = GetFieldValByLabelB("Other:")
    otherStaff2 = Trim(wsInput.Range("D29").value)
    otherStaff3 = Trim(wsInput.Range("E29").value)
    
    ' --- Comma validation for all fields being written to file ---
    If Not FormUtils.ValidateNoCommas(quickCode, "Quick Code") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(productName, "Product") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(pmiCode, "PMI Number") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(pmiVersion, "PMI Version") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(batchNumber, "Batch Number") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(startRoom, "Start Room") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(finishRoom, "Finish Room") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(CStr(plannedStartDate), "Planned Start Date") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(CStr(actualStartDate), "Actual Start Date") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(CStr(actualFinishDate), "Actual Finish Date") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(CStr(plannedReleaseQty), "Planned Release Qty") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(CStr(sampleQty), "Sample Qty") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(notesText, "Notes") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(prepStaff1, "Production Preparation Staff 1") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(prepStaff2, "Production Preparation Staff 2") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(prepStaff3, "Production Preparation Staff 3") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(makeStaff1, "Make Staff 1") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(makeStaff2, "Make Staff 2") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(makeStaff3, "Make Staff 3") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(checkStaff1, "Checking Staff 1") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(checkStaff2, "Checking Staff 2") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(checkStaff3, "Checking Staff 3") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(fillStaff1, "Fill Staff 1") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(fillStaff2, "Fill Staff 2") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(fillStaff3, "Fill Staff 3") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(labelStaff1, "Label Staff 1") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(labelStaff2, "Label Staff 2") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(labelStaff3, "Label Staff 3") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(packStaff1, "Pack Staff 1") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(packStaff2, "Pack Staff 2") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(packStaff3, "Pack Staff 3") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(otherStaff1, "Other Staff 1") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(otherStaff2, "Other Staff 2") Then GoTo CleanExit
    If Not FormUtils.ValidateNoCommas(otherStaff3, "Other Staff 3") Then GoTo CleanExit
    
    If IsNumeric(sampleQty) And sampleQty <> "" Then
        totalQty = CLng(plannedReleaseQty) + CLng(sampleQty)
    Else
        totalQty = CLng(plannedReleaseQty)
    End If
    
    If quickCode = "" Then
        response = MsgBox("Quick Code is empty. Click Yes to continue anyway, or No to cancel.", vbYesNo + vbExclamation, "Save_ProdOutput")
        If response = vbNo Then
            GoTo CleanExit
        End If
    End If
    
    If batchNumber = "" Then
        MsgBox "Batch Number is required.", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    If Not batchNumber Like "[A-Za-z]######" Then
        MsgBox "Batch Number must be in the format: 1 letter followed by 6 digits (e.g., H123456).", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    batchValidationResult = ValidateAndHandleBatchStatusProductionOutput(batchNumber, "Save_ProdOutput", requiresStatusUpdate)
    
    If batchValidationResult <> bsrAllow Then
        GoTo CleanExit
    End If
    
    If startRoom = "" Then
        MsgBox "Start Room is required.", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    If plannedStartDate = "" Or Not IsDate(plannedStartDate) Then
        MsgBox "A valid Planned Start Date is required.", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    If actualStartDate = "" Or Not IsDate(actualStartDate) Then
        MsgBox "A valid Actual Start Date is required.", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    If actualFinishDate = "" Or Not IsDate(actualFinishDate) Then
        MsgBox "A valid Actual Finish Date is required.", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    If plannedReleaseQty = "" Or Not IsNumeric(plannedReleaseQty) Then
        MsgBox "Planned Release Quantity must be a number.", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    If sampleQty <> "" And Not IsNumeric(sampleQty) Then
        MsgBox "Sample Quantity must be a number or left blank.", vbExclamation, "Save_ProdOutput"
        GoTo CleanExit
    End If
    
    If BatchNumberExists(batchNumber) Then
        If notesText = "" Then
            MsgBox "Notes are mandatory when updating an existing production output record.", vbExclamation, "Save_ProdOutput"
            GoTo CleanExit
        End If
        
        response = MsgBox("Production output for batch already entered. Click Yes to accept the entered details, or No to change.", vbYesNo + vbExclamation, "Save_ProdOutput")
        If response = vbNo Then
            GoTo CleanExit
        End If
    End If
    
    Set prodMasterWs = ThisWorkbook.Sheets("FG_Master")
    Set prodMasterRange = prodMasterWs.UsedRange
    
    colA = 1
    colB = 2
    colC = 3
    colJ = 10
    
    productType = ""
    productCategory = ""
    
    For Each lookupRow In prodMasterRange.Rows
        If lookupRow.Cells(1, colA).value = pmiCode And lookupRow.Cells(1, colB).value = pmiVersion Then
            productType = lookupRow.Cells(1, colC).value
            productCategory = lookupRow.Cells(1, colJ).value
            Exit For
        End If
    Next lookupRow
    
    outputData(1) = DateToANSI(Format$(actualStartDate, "dd/mm/yyyy"))
    outputData(2) = DateToANSI(Format$(plannedStartDate, "dd/mm/yyyy"))
    outputData(3) = productName
    outputData(4) = quickCode
    outputData(5) = pmiCode
    outputData(6) = productType
    outputData(7) = productCategory
    outputData(8) = startRoom
    outputData(9) = finishRoom
    outputData(10) = batchNumber
    outputData(11) = CStr(plannedReleaseQty)
    outputData(12) = prepStaff1
    outputData(13) = prepStaff2
    outputData(14) = prepStaff3
    outputData(15) = makeStaff1
    outputData(16) = makeStaff2
    outputData(17) = makeStaff3
    outputData(18) = checkStaff1
    outputData(19) = checkStaff2
    outputData(20) = checkStaff3
    outputData(21) = fillStaff1
    outputData(22) = fillStaff2
    outputData(23) = fillStaff3
    outputData(24) = labelStaff1
    outputData(25) = labelStaff2
    outputData(26) = labelStaff3
    outputData(27) = packStaff1
    outputData(28) = packStaff2
    outputData(29) = packStaff3
    outputData(30) = otherStaff1
    outputData(31) = otherStaff2
    outputData(32) = otherStaff3
    outputData(33) = notesText
    outputData(34) = DateToANSI(Format$(actualFinishDate, "dd/mm/yyyy"))
    outputData(35) = Environ("USERNAME")
    outputData(36) = DateToANSI(Format$(actualFinishDate, "dd/mm/yyyy")) & " " & Format$(Now, "hh:mm:ss")
    outputData(37) = CStr(sampleQty)
    outputData(38) = CStr(totalQty)
    
    If requiresStatusUpdate Then
        batchTransID = GetNextFinishedGoodsBatchTransID()
        outputData(39) = batchTransID
        
        batchUpdateSuccess = UpdateFinishedGoodsBatchRecord( _
            batchNumber:=batchNumber, _
            batchStatus:=FGB_STATUS_COMPLETE, _
            dateUpdated:=DateToANSI(Format$(Date, "dd/mm/yyyy")), _
            updatedBy:=Environ$("USERNAME"), _
            updateComments:="Production Output Saved", _
            auditTransID:=batchTransID _
        )
        
        If Not batchUpdateSuccess Then
            MsgBox "Warning: Production output saved but batch status update failed.", vbExclamation, "Batch Update Warning"
        End If
    Else
        outputData(39) = ""
    End If
    
    WriteProductionOutputRecord outputData
    
    MsgBox "Production output record saved successfully.", vbInformation, "Save_ProdOutput"
    
    wsInput.Range("C5:C13, C17:C18, C23:E29").ClearContents
    wsInput.Range("C19").value = ""

    FGBatchDetails
    
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Save_ProdOutput"
    
CleanExit:
    Set lookupRow = Nothing
    Set prodMasterRange = Nothing
    Set prodMasterWs = Nothing
    Set rngProdLookup = Nothing
    Set wsInput = Nothing
    AppStateUtils.RestoreAppState
    FormUtils.ProtectActiveSheet
    Exit Sub
End Sub

Private Function BatchNumberExists(ByVal batchNumber As String) As Boolean
    Dim conn As Object
    Dim rs As Object
    Dim folderPath As String
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    BatchNumberExists = False
    
    folderPath = GetWorkbookTablesPath()
    Set conn = OpenADOConnection(folderPath)
    
    If conn Is Nothing Then
        GoTo CleanExit
    End If
    
    sql = "SELECT COUNT(*) FROM [" & FILE_PRODUCTION_OUTPUT & "] " & _
          "WHERE [" & OUTPUT_HEADER_BATCH_NUMBER & "] = '" & Replace(batchNumber, "'", "''") & "'"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        BatchNumberExists = (rs.Fields(0).value > 0)
    End If
    
    GoTo CleanExit
    
ErrHandler:
    BatchNumberExists = False
    
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
End Function
