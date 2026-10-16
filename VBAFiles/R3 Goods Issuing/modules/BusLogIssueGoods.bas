Attribute VB_Name = "BusLogIssueGoods"
' ========================================================================
' GOODS ISSUE TRANSACTION PROCESSING
' ========================================================================
' Validates QC status, retest dates, balance, converts UoM, and writes to
' RM_Trans.csv. Includes duplicate QC Number detection for batch.
' ========================================================================

Option Explicit

' True when the last SaveGoodsIssue call wrote a batch status change to FinishedGoodsBatchNumberMasterFile.csv
Public goodsIssueChangedBatchStatus As Boolean

' Main procedure to save goods issue (consumption) transaction
' Validates QC status, retest dates, balance, converts UoM, and writes to RM_Trans.csv
Public Sub SaveGoodsIssue()
    On Error GoTo ErrorHandler
    goodsIssueChangedBatchStatus = False
    
    Dim batchNumber As String
    Dim rmCode As String
    Dim stdName As String
    Dim qcNumber As String
    Dim issueQtyStr As String
    Dim issueUoM As String
    Dim overageFlag As String
    Dim overageQtyStr As String
    Dim notesComments As String
    Dim measuredBy As String
    Dim masterUoM As String
    Dim qcStatus As String
    Dim nilStockFlag As String
    Dim qcCSVPath As String
    Dim qcChangesPath As String
    Dim qcTransID As String
    Dim fileContent As String
    Dim line As String
    Dim issueQty As Double
    Dim overageQty As Double
    Dim convertedIssueQty As Double
    Dim convertedOverageQty As Double
    Dim totalQty As Double
    Dim currentBalance As Double
    Dim newBalance As Double
    Dim wtPerML As Double
    Dim retestDate As Variant
    Dim firstReleaseDate As Variant
    Dim retestDateVal As Date
    Dim daysDiff As Long
    Dim qcRow As Long
    Dim rmRow As Long
    Dim colQCNumber As Long
    Dim colQCStatus As Long
    Dim colNilStock As Long
    Dim colRetestDate As Long
    Dim colQCBalance As Long
    Dim colRMCode As Long
    Dim colMasterUoM As Long
    Dim colWtPerML As Long
    Dim colFirstReleaseDate As Long
    Dim i As Long
    Dim r As Long
    Dim lastRow As Long
    Dim idxQC As Long
    Dim idxQCBal As Long
    Dim idxTransDate As Long
    Dim idxTransUser As Long
    Dim idxNilStockFlag As Long
    Dim idxNilStockFlagDate As Long
    Dim idxNilStockFlagBy As Long
    Dim idxQCTransID As Long
    Dim idxChangeReason As Long
    Dim found As Boolean
    Dim nilStockUpdate As Boolean
    Dim requiresStatusUpdate As Boolean
    Dim batchValidationResult As BatchStatusResult
    Dim wsQC As Worksheet
    Dim wsRM As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim lines() As String
    Dim lineParts() As String
    Dim response As VbMsgBoxResult
    
    ' --- 1. Get Mandatory Fields ---
    batchNumber = Trim$(GetFieldValByLabelC("Batch Number:"))
    rmCode = Trim$(GetFieldValByLabelC("RM Code:"))
    stdName = Trim$(GetFieldValByLabelC("Standard Name:"))
    qcNumber = Trim$(GetFieldValByLabelC("QC Number:"))
    issueQtyStr = Trim$(GetFieldValByLabelC("Issue Quantity:"))
    issueUoM = Trim$(GetFieldValByLabelC("Issue UoM:"))
    overageFlag = Trim$(GetFieldValByLabelC("Overage:"))
    overageQtyStr = Trim$(GetFieldValByLabelC("Overage Quantity:"))
    notesComments = Trim$(GetFieldValByLabelC("Notes/Comments:"))
    measuredBy = Trim$(GetFieldValByLabelC("Measured By:"))
    
    ' Validate that critical text fields do not contain commas (which would break CSV formatting)
    If Not ValidateNoCommas(notesComments, "Notes/Comments") Then GoTo CleanExit
    If Not ValidateNoCommas(measuredBy, "Measured By") Then GoTo CleanExit
    
    ' Set default for overage flag
    If overageFlag = "" Then overageFlag = "No"
    
    ' Validate all mandatory fields are present
    If batchNumber = "" Or qcNumber = "" Or issueQtyStr = "" Or issueUoM = "" Or measuredBy = "" Then
        MsgBox "Please complete all mandatory fields:" & vbCrLf & vbCrLf & _
               "- Batch Number" & vbCrLf & _
               "- QC Number" & vbCrLf & _
               "- Issue Quantity" & vbCrLf & _
               "- Issue UoM" & vbCrLf & _
               "- Measured By", vbExclamation, "Missing Information"
        GoTo CleanExit
    End If
    
    ' --- 1a. Validate Batch Status ---
    batchValidationResult = ValidateAndHandleBatchStatusGoodsIssue(batchNumber, "SaveGoodsIssue", requiresStatusUpdate)
    
    If batchValidationResult <> bsrAllow Then
        GoTo CleanExit
    End If
    
    ' Validate Issue Quantity is numeric
    If Not IsNumeric(issueQtyStr) Then
        MsgBox "Issue Quantity must be a valid number.", vbExclamation
        GoTo CleanExit
    End If
    issueQty = CDbl(issueQtyStr)
    
    If issueQty <= 0 Then
        MsgBox "Issue Quantity must be greater than zero.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Validate Overage Quantity
    overageQty = 0
    If UCase$(overageFlag) = "YES" Then
        If overageQtyStr = "" Or Not IsNumeric(overageQtyStr) Then
            MsgBox "Overage Quantity must be provided and numeric when Overage is Yes.", vbExclamation
            GoTo CleanExit
        End If
        overageQty = CDbl(overageQtyStr)
        If overageQty <= 0 Then
            MsgBox "Overage Quantity must be greater than zero when Overage is Yes.", vbExclamation
            GoTo CleanExit
        End If
    Else
        If overageQtyStr <> "" And IsNumeric(overageQtyStr) Then
            If CDbl(overageQtyStr) > 0 Then
                MsgBox "Overage Quantity has been entered but Overage is set to No." & vbCrLf & _
                       "Please set Overage to Yes or clear the Overage Quantity.", vbExclamation
                GoTo CleanExit
            End If
        End If
    End If
    
    ' --- 2. Check if QC Number already issued to batch ---
    If CheckQCNumberIssuedToBatch(batchNumber, qcNumber) Then
        response = MsgBox("QC Number " & qcNumber & " has already been issued to Batch " & batchNumber & "." & vbCrLf & vbCrLf & _
                         "Do you want to continue?", vbYesNo + vbExclamation + vbDefaultButton2, "Duplicate QC Issue")
        If response = vbNo Then
            GoTo CleanExit
        End If
    End If
    
    ' --- 3. Validate QC Number (from QC_Number worksheet) ---
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    ' Find column indices
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colQCStatus = FindHeaderCol(wsQC, 1, HEADER_QC_STATUS_FLAG)
    colNilStock = FindHeaderCol(wsQC, 1, HEADER_NIL_STOCK_FLAG)
    colRetestDate = FindHeaderCol(wsQC, 1, HEADER_LAST_RETEST_DATE)
    colQCBalance = FindHeaderCol(wsQC, 1, HEADER_QC_BALANCE)
    
    If colQCNumber = 0 Or colQCStatus = 0 Or colNilStock = 0 Or colRetestDate = 0 Or colQCBalance = 0 Then
        MsgBox "Required columns not found in QC_Number worksheet.", vbCritical
        GoTo CleanExit
    End If
    
    ' Find QC Number row
    qcRow = 0
    lastRow = wsQC.Cells(wsQC.Rows.Count, colQCNumber).End(xlUp).Row
    
    For r = 2 To lastRow
        If Trim$(CStr(wsQC.Cells(r, colQCNumber).value)) = qcNumber Then
            qcRow = r
            Exit For
        End If
    Next r
    
    If qcRow = 0 Then
        MsgBox "QC Number " & qcNumber & " not found in database.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Get QC Status
    qcStatus = Trim$(CStr(wsQC.Cells(qcRow, colQCStatus).value))
    If qcStatus = QC_STATUS_RELEASED Then
        ' Released status is always valid
    ElseIf qcStatus = QC_STATUS_SAMPLED Then

        ' Sampled status requires First Release Date
        colFirstReleaseDate = FindHeaderCol(wsQC, 1, HEADER_FIRST_RELEASE_DATE)
        If colFirstReleaseDate = 0 Then
            MsgBox "Required column 'First_Release_Date' not found in QC_Number worksheet.", vbCritical
            GoTo CleanExit
        End If
        
        firstReleaseDate = wsQC.Cells(qcRow, colFirstReleaseDate).value
        
        If IsEmpty(firstReleaseDate) Or Not IsDate(firstReleaseDate) Then
            MsgBox "QC Number " & qcNumber & " has Sampled status but no First Release Date." & vbCrLf & _
                   "Cannot issue from this lot.", vbExclamation, "Invalid QC Status"
            GoTo CleanExit
        End If
    Else
        MsgBox "QC Number " & qcNumber & " must be Released or Sampled (with First Release Date)." & vbCrLf & _
               "Current status: " & qcStatus, vbExclamation, "Invalid QC Status"
        GoTo CleanExit
    End If
    
    ' Check Nil Stock Flag
    nilStockFlag = Trim$(CStr(wsQC.Cells(qcRow, colNilStock).value))
    If nilStockFlag = "" Then nilStockFlag = "N"
    If UCase$(nilStockFlag) = "Y" Then
        MsgBox "QC Number " & qcNumber & " is flagged as Nil Stock." & vbCrLf & _
               "Cannot issue from this lot.", vbExclamation, "Nil Stock"
        GoTo CleanExit
    End If
    
    ' Check Retest Date
    retestDate = wsQC.Cells(qcRow, colRetestDate).value
    If IsEmpty(retestDate) Or Not IsDate(retestDate) Then
        MsgBox "QC Number " & qcNumber & " has no valid retest date." & vbCrLf & _
               "Cannot issue from this lot.", vbExclamation, "Invalid Retest Date"
        GoTo CleanExit
    End If
    
    retestDateVal = CDate(retestDate)
    
    ' Check if retest date has lapsed
    If retestDateVal < Date Then
        MsgBox "QC Number " & qcNumber & " retest date has lapsed." & vbCrLf & _
               "Retest Date: " & Format$(retestDateVal, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
               "Cannot issue from this lot.", vbExclamation, "Retest Date Lapsed"
        GoTo CleanExit
    End If
    
    ' Warn if retest date is within 30 days
    daysDiff = retestDateVal - Date
    If daysDiff <= 30 Then
        response = MsgBox("QC Number " & qcNumber & " retest date expires within 30 days." & vbCrLf & _
                         "Retest Date: " & Format$(retestDateVal, "dd/mm/yyyy") & vbCrLf & _
                         "Days remaining: " & daysDiff & vbCrLf & vbCrLf & _
                         "Do you want to continue?", vbYesNo + vbExclamation, "Retest Date Warning")
        If response = vbNo Then
            GoTo CleanExit
        End If
    End If
    
    ' --- 4. Get Master UoM & Convert Quantity ---
    Set wsRM = ThisWorkbook.Worksheets("RM_Master")
    
    colRMCode = FindHeaderCol(wsRM, 1, RMM_HEADER_RM_CODE)
    colMasterUoM = FindHeaderCol(wsRM, 1, RMM_HEADER_MASTER_UOM)
    colWtPerML = FindHeaderCol(wsRM, 1, RMM_HEADER_WT_PER_ML)
    
    If colRMCode = 0 Or colMasterUoM = 0 Or colWtPerML = 0 Then
        MsgBox "Required columns not found in RM_Master worksheet.", vbCritical
        GoTo CleanExit
    End If
    
    ' Find RM Code row
    rmRow = 0
    lastRow = wsRM.Cells(wsRM.Rows.Count, colRMCode).End(xlUp).Row
    
    For r = 2 To lastRow
        If Trim$(CStr(wsRM.Cells(r, colRMCode).value)) = rmCode Then
            rmRow = r
            Exit For
        End If
    Next r
    
    If rmRow = 0 Then
        MsgBox "RM Code " & rmCode & " not found in RM Master.", vbCritical
        GoTo CleanExit
    End If
    
    masterUoM = Trim$(CStr(wsRM.Cells(rmRow, colMasterUoM).value))
    
    wtPerML = 0
    On Error Resume Next
    wtPerML = CDbl(wsRM.Cells(rmRow, colWtPerML).value)
    On Error GoTo ErrorHandler
    
    ' Convert quantity if UoM differs
    If UCase$(issueUoM) <> UCase$(masterUoM) Then
        convertedIssueQty = ConvertToMasterUoM(issueQty, issueUoM, masterUoM, wtPerML)
        If convertedIssueQty <= 0 Then
            MsgBox "Unable to convert " & issueUoM & " to " & masterUoM & " for RM Code " & rmCode & vbCrLf & _
                   "Wt_per_mL: " & wtPerML, vbExclamation
            GoTo CleanExit
        End If
        
        If overageQty > 0 Then
            convertedOverageQty = ConvertToMasterUoM(overageQty, issueUoM, masterUoM, wtPerML)
            If convertedOverageQty <= 0 Then
                MsgBox "Unable to convert Overage " & issueUoM & " to " & masterUoM & " for RM Code " & rmCode & vbCrLf & _
                       "Wt_per_mL: " & wtPerML, vbExclamation
                GoTo CleanExit
            End If
        Else
            convertedOverageQty = 0
        End If
    Else
        convertedIssueQty = issueQty
        convertedOverageQty = overageQty
    End If
    
    totalQty = convertedIssueQty + convertedOverageQty
    
    ' --- 5. Check QC Balance ---
    currentBalance = GetQCBalanceFromRMTrans(qcNumber)

    
    If Round(currentBalance, 3) < Round(totalQty, 3) Then
        MsgBox "Insufficient QC Balance for QC Number " & qcNumber & vbCrLf & vbCrLf & _
               "Available: " & Format$(currentBalance, "0.000") & " " & masterUoM & vbCrLf & _
               "Required: " & Format$(totalQty, "0.000") & " " & masterUoM, _
               vbExclamation, "Insufficient Balance"
        GoTo CleanExit
    End If
    
    ' --- 6. Determine if Nil Stock update is needed ---
    newBalance = currentBalance - totalQty
    nilStockUpdate = False
    qcTransID = ""
    
    If Round(newBalance, 6) = 0 Then
        response = MsgBox("Issuing this quantity will consume the remaining balance for QC Number " & qcNumber & "." & vbCrLf & vbCrLf & _
                      "This will mark the QC lot as Nil Stock. Do you want to continue?", vbYesNo + vbExclamation, "Consume Remaining Balance")
        If response = vbNo Then
            GoTo CleanExit
        End If
        nilStockUpdate = True
        qcTransID = GetNextQCTransactionID()
    End If
    
    ' --- 7. Write Transactions to RM_Trans.csv ---
    Dim issueNote As String
    
    If UCase$(issueUoM) <> UCase$(masterUoM) Then
        issueNote = "Converted measured quantity of " & Format$(issueQty, "0.000") & " " & issueUoM & _
                    " converted to " & Format$(convertedIssueQty, "0.000") & " " & masterUoM
        If notesComments <> "" Then
            issueNote = issueNote & ". " & notesComments
        End If
    Else
        issueNote = notesComments
    End If
    
    Call WriteRMConsumptionTransaction( _
        transType:=TRANS_TYPE_CONSUMED, _
        qcNumber:=qcNumber, _
        qcStat:=QC_STATUS_RELEASED, _
        fgBatch:=batchNumber, _
        fgBatchStatus:="", _
        quantity:=-convertedIssueQty, _
        masterUoM:=masterUoM, _
        adjNote:=issueNote, _
        piTransID:="NA", _
        qcTransID:=IIf(nilStockUpdate, qcTransID, "NA"), _
        measuredBy:=measuredBy _
    )
    
    If convertedOverageQty > 0 Then
        Dim overageNote As String
        
        If UCase$(issueUoM) <> UCase$(masterUoM) Then
            overageNote = "OVERAGE: Converted measured quantity of " & Format$(overageQty, "0.000") & " " & issueUoM & _
                         " converted to " & Format$(convertedOverageQty, "0.000") & " " & masterUoM
            If notesComments <> "" Then
                overageNote = overageNote & ". " & notesComments
            End If
        Else
            overageNote = "OVERAGE: " & notesComments
        End If
        
        Call WriteRMConsumptionTransaction( _
            transType:=TRANS_TYPE_CONSUMED, _
            qcNumber:=qcNumber, _
            qcStat:=QC_STATUS_RELEASED, _
            fgBatch:=batchNumber, _
            fgBatchStatus:="", _
            quantity:=-convertedOverageQty, _
            masterUoM:=masterUoM, _
            adjNote:=overageNote, _
            piTransID:="NA", _
            qcTransID:=IIf(nilStockUpdate, qcTransID, "NA"), _
            measuredBy:=measuredBy _
        )
    End If
    
    ' --- 8. Update QC_Number.csv ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    qcCSVPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    Set ts = fso.OpenTextFile(qcCSVPath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    
    lines = Split(fileContent, vbCrLf)
    found = False
    
    idxQC = GetHeaderIndex0(HEADER_QC_NUMBER)
    idxQCBal = GetHeaderIndex0(HEADER_QC_BALANCE)
    idxTransDate = GetHeaderIndex0(HEADER_TRANS_DATE)
    idxTransUser = GetHeaderIndex0(HEADER_TRANS_USER)
    idxNilStockFlag = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG)
    idxNilStockFlagDate = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_DATE)
    idxNilStockFlagBy = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_BY)
    idxQCTransID = GetHeaderIndex0(HEADER_QC_TRANS_ID)
    idxChangeReason = GetHeaderIndex0(HEADER_CHANGE_REASON)
    
    For i = 1 To UBound(lines)
        line = lines(i)
        If Trim(line) <> "" Then
            lineParts = Split(line, ",")
            
            If UBound(lineParts) >= idxQC And _
                LCase$(Trim$(Replace(lineParts(idxQC), """", ""))) = LCase$(Trim$(qcNumber)) Then
                
                lineParts(idxQCBal) = """" & CStr(Round(newBalance, 3)) & """"
                lineParts(idxChangeReason) = """" & "Goods Issue" & """"
                
                If idxTransDate > -1 Then lineParts(idxTransDate) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"
                If idxTransUser > -1 Then lineParts(idxTransUser) = """" & Environ$("Username") & """"
                
                If nilStockUpdate Then
                    If idxNilStockFlag > -1 Then lineParts(idxNilStockFlag) = """Y"""
                    If idxNilStockFlagDate > -1 Then lineParts(idxNilStockFlagDate) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"
                    If idxNilStockFlagBy > -1 Then lineParts(idxNilStockFlagBy) = """" & Environ$("Username") & """"
                    If idxQCTransID > -1 Then lineParts(idxQCTransID) = """" & qcTransID & """"
                End If
                
                lines(i) = Join(lineParts, ",")
                found = True
                Exit For
            End If
        End If
    Next i
    
    If Not found Then
        MsgBox "QC Number '" & qcNumber & "' not found in QC_Number.csv. No entry was updated.", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- 9. Write to QC_Number_Changes.csv FIRST (AUDIT TRAIL) ---
    If nilStockUpdate Then
        Dim auditPath As String
        auditPath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
        Set ts = fso.OpenTextFile(auditPath, 8, True)
        ts.WriteLine lines(i)
        ts.Close
        Set ts = Nothing
    End If
    
    ' --- 10. Now overwrite the QC_Number.csv file with the updated content ---
    Set ts = fso.OpenTextFile(qcCSVPath, 2, False)
    ts.Write Join(lines, vbCrLf)
    ts.Close
    Set ts = Nothing
    
    ' --- 11. Update Finished Goods Batch Status (only if required) ---
    If requiresStatusUpdate Then
        Dim batchUpdateSuccess As Boolean
        batchUpdateSuccess = UpdateFinishedGoodsBatchRecord( _
            batchNumber:=batchNumber, _
            batchStatus:=FGB_STATUS_INPROGRESS, _
            dateUpdated:=DateToANSI(Format$(Date, "dd/mm/yyyy")), _
            updatedBy:=Environ$("USERNAME"), _
            updateComments:="Goods Issue - QC Number: " & qcNumber _
        )
        
        goodsIssueChangedBatchStatus = batchUpdateSuccess
        If Not batchUpdateSuccess Then
            MsgBox "Warning: Goods Issue completed but batch status update failed.", vbExclamation, "Batch Update Warning"
        End If
    End If
    
    MsgBox "Raw Material has been issued." & vbCrLf & vbCrLf & _
           "QC Number: " & qcNumber & vbCrLf & _
           "Issue Quantity: " & Format$(convertedIssueQty, "0.000") & " " & masterUoM & vbCrLf & _
           IIf(convertedOverageQty > 0, "Overage Quantity: " & Format$(convertedOverageQty, "0.000") & " " & masterUoM & vbCrLf, "") & _
           "New Balance: " & Format$(newBalance, "0.000") & " " & masterUoM, _
           vbInformation, "Issue Complete"
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error saving goods issue: " & Err.Description, vbCritical, "SaveGoodsIssue"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set wsQC = Nothing
    Set wsRM = Nothing
    Exit Sub
End Sub

Private Function CheckQCNumberIssuedToBatch(ByVal batchNumber As String, ByVal qcNumber As String) As Boolean
    Dim conn As Object
    Dim rs As Object
    Dim fso As Object
    Dim tempFilePath As String
    Dim tempFolder As String
    Dim sqlQuery As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    tempFilePath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempFilePath = "" Then
        CheckQCNumberIssuedToBatch = False
        GoTo CleanExit
    End If
    
    tempFolder = fso.GetParentFolderName(tempFilePath)
    
    Set conn = OpenADOConnection(tempFolder)
    
    If conn Is Nothing Then
        CheckQCNumberIssuedToBatch = False
        GoTo CleanExit
    End If
    
    sqlQuery = "SELECT COUNT(*) AS RecordCount " & _
               "FROM [" & fso.GetFileName(tempFilePath) & "] " & _
               "WHERE " & RMT_HEADER_QC_NUMBER & " = '" & Replace(qcNumber, "'", "''") & "' " & _
               "AND " & RMT_HEADER_FG_BATCH & " = '" & Replace(batchNumber, "'", "''") & "'"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn
    
    If Not rs.EOF Then
        Dim recordCount As Long
        recordCount = 0
        
        On Error Resume Next
        recordCount = CLng(rs.Fields("RecordCount").value)
        On Error GoTo ErrorHandler
        
        CheckQCNumberIssuedToBatch = (recordCount > 0)
    Else
        CheckQCNumberIssuedToBatch = False
    End If
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    GoTo CleanExit
    
ErrorHandler:
    CheckQCNumberIssuedToBatch = False
    
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
    Exit Function
End Function
