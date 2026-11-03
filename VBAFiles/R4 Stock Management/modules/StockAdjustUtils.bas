Attribute VB_Name = "StockAdjustUtils"
Option Explicit

' ========================================================================
' CENTRALIZED STOCK UPDATE MODULE
' ========================================================================
' Orchestrates stock movements by calling existing utility functions
' and ensuring consistent business logic across all stock movement types
' ========================================================================

Public Type StockUpdateResult
    Success As Boolean
    ErrorMessage As String
    PreviousBalance As Double
    newBalance As Double
    transactionQty As Double
    adjustmentType As String
    NilStockAffected As Boolean
    qcTransID As String
    masterUoM As String
End Type

' ========================================================================
' Main entry point for all stock updates
' ========================================================================
Public Function ProcessStockUpdate( _
    ByVal qcNumber As String, _
    ByVal inputQty As Double, _
    ByVal inputUoM As String, _
    ByVal reason As String, _
    Optional ByVal adjustmentType As String = "", _
    Optional ByVal confirmNilStock As Boolean = True, _
    Optional ByVal fgBatch As String = "", _
    Optional ByVal fgBatchStatus As String = "", _
    Optional ByVal allowZeroQty As Boolean = False, _
    Optional ByVal piNumber As String = "") As StockUpdateResult
    
    Dim result As StockUpdateResult
    Dim rmCode As String
    Dim masterUoM As String
    Dim convertedQty As Double
    Dim currentBalance As Double
    Dim newBalance As Double
    Dim transactionQty As Double
    Dim wtPerML As Double
    Dim nilStockUpdate As Boolean
    Dim qcTransID As String
    Dim currentNilStockFlag As String
    Dim currentQCStatus As String
    Dim transType As String
    Dim logStocktake As Boolean
    Dim stocktakeLogPath As String
    Dim tsLog As Object

    On Error GoTo ErrorHandler
    
    result.Success = False
    result.ErrorMessage = ""
    result.NilStockAffected = False
    result.qcTransID = ""
    
    ' --- STEP 1: Validate Input Parameters ---
    If Trim$(qcNumber) = "" Then
        result.ErrorMessage = "QC Number cannot be blank."
        GoTo CleanExit
    End If
    
    If inputQty < 0 Then
        result.ErrorMessage = "Input quantity cannot be negative."
        GoTo CleanExit
    End If
    
    If inputQty = 0 And Not allowZeroQty Then
        result.ErrorMessage = "Input quantity must be greater than zero."
        GoTo CleanExit
    End If
    
    If Trim$(inputUoM) = "" Then
        result.ErrorMessage = "Input UoM cannot be blank."
        GoTo CleanExit
    End If
    
    ' --- STEP 2: Validate QC Number exists and get RM Code, Nil Stock Flag, and QC Status ---
    If Not GetRMCodeFromQCNumber(qcNumber, rmCode, currentNilStockFlag, currentQCStatus) Then
        result.ErrorMessage = "QC Number '" & qcNumber & "' not found in database."
        GoTo CleanExit
    End If
    
    If Trim$(piNumber) = "" Then
        If UCase$(currentQCStatus) = UCase$(QC_STATUS_CANCELLED) Or UCase$(currentQCStatus) = UCase$(QC_STATUS_REJECTED) Then
            result.ErrorMessage = "QC Number '" & qcNumber & "' has a status of '" & currentQCStatus & "'." & vbCrLf & vbCrLf & _
                   "Stock adjustments cannot be recorded for Cancelled or Rejected QC Numbers."
            GoTo CleanExit
        End If
    End If
    
    ' --- STEP 3: Get Master UoM and conversion factor ---
    If Not GetMasterUoMAndConversionFactor(rmCode, masterUoM, wtPerML) Then
        result.ErrorMessage = "Unable to retrieve master UoM for RM Code '" & rmCode & "'."
        GoTo CleanExit
    End If
    
    result.masterUoM = masterUoM
    
    ' --- STEP 4: Convert input quantity to master UoM ---
    ' Handle zero qty specially (no conversion needed, just use zero)
    If inputQty = 0 Then
        convertedQty = 0
    Else
        convertedQty = ConvertToMasterUoM(inputQty, inputUoM, masterUoM, wtPerML)
        If convertedQty <= 0 Then
            result.ErrorMessage = "Unable to convert " & inputUoM & " to " & masterUoM & " for RM Code " & rmCode
            GoTo CleanExit
        End If
    End If
    
    ' --- STEP 5: Get current balance from RM_Trans.csv ---
    currentBalance = GetQCBalanceADO(qcNumber)
    
    result.PreviousBalance = currentBalance
    
    ' --- STEP 6: Determine Adjustment Type (if not provided) ---
    If Trim$(adjustmentType) = "" Then
        ' Auto-determine based on input vs current balance (for stocktake scenarios)
        If Round(convertedQty, 3) > Round(currentBalance, 3) Then
            adjustmentType = "WRITE-ON"
            transactionQty = convertedQty - currentBalance
        ElseIf Round(convertedQty, 3) < Round(currentBalance, 3) Then
            adjustmentType = "WRITE-OFF"
            transactionQty = -(currentBalance - convertedQty)
        Else
            ' No change in balance
            result.ErrorMessage = "Input quantity matches current balance. No adjustment needed."
            GoTo CleanExit
        End If
    Else
        ' Manual adjustment type provided (from Stock Adjustment form)
        adjustmentType = UCase$(Trim$(adjustmentType))
        If adjustmentType = "WRITE-OFF" Then
            transactionQty = -convertedQty
        ElseIf adjustmentType = "WRITE-ON" Then
            transactionQty = convertedQty
        Else
            result.ErrorMessage = "Invalid adjustment type. Must be 'WRITE-ON' or 'WRITE-OFF'."
            GoTo CleanExit
        End If
    End If
    
    result.adjustmentType = adjustmentType
    result.transactionQty = transactionQty
    
    ' --- STEP 7: Validate balance sufficiency for write-offs ---
    If adjustmentType = "WRITE-OFF" Then
        If currentBalance < Abs(transactionQty) Then
            result.ErrorMessage = "Insufficient QC Balance for QC Number " & qcNumber & vbCrLf & _
                                "Available: " & Format$(currentBalance, "0.000") & " " & masterUoM & vbCrLf & _
                                "Required: " & Format$(Abs(transactionQty), "0.000") & " " & masterUoM
            GoTo CleanExit
        End If
    End If
    
    ' --- STEP 8: Calculate new balance ---
    newBalance = currentBalance + transactionQty
    result.newBalance = newBalance
    
    ' --- STEP 9: Generate QC_Trans_ID ---
    qcTransID = GetNextQCTransactionID()
    result.qcTransID = qcTransID
    
    ' --- STEP 10: Write Transaction to RM_Trans.csv ---
    If adjustmentType = "WRITE-OFF" Then
        transType = TRANS_TYPE_WRITE_OFF
    Else
        transType = TRANS_TYPE_WRITE_ON
    End If
            
    Call WriteRMConsumptionTransaction( _
        transType:=transType, _
        qcNumber:=qcNumber, _
        qcStat:=currentQCStatus, _
        fgBatch:=fgBatch, _
        fgBatchStatus:=fgBatchStatus, _
        quantity:=transactionQty, _
        masterUoM:=masterUoM, _
        adjNote:=reason, _
        qcTransID:=qcTransID, _
        piTransID:=IIf(piNumber <> "", piNumber, "NA"))
        
    ' --- STEP 11: Update QC_Number.csv ---
    Dim updateResult As Boolean
    Dim changeReasonText As String

    If adjustmentType = "WRITE-OFF" And Round(newBalance, 6) = 0 Then
        ' Exhausting all stock - mark as nil stock
        If confirmNilStock Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("This transaction will consume the remaining balance for QC Number " & qcNumber & "." & vbCrLf & vbCrLf & _
                         "This will mark the QC lot as Nil Stock. Do you want to continue?", _
                         vbYesNo + vbExclamation, "Consume Remaining Balance")
            If resp = vbNo Then
                result.ErrorMessage = "Transaction cancelled by user."
                GoTo CleanExit
            End If
        End If
        changeReasonText = piNumber & " " & reason & " - Nil Stock"
        updateResult = UpdateQCNumberRecord( _
            qcNumber:=qcNumber, _
            qcBalance:=newBalance, _
            nilStockFlag:="Y", _
            nilStockFlagDate:=Format$(Date, "dd/mm/yyyy"), _
            nilStockFlagBy:=Environ$("Username"), _
            changeReason:=changeReasonText, _
            qcTransID:=qcTransID)
    ElseIf adjustmentType = "WRITE-ON" And UCase$(currentNilStockFlag) = "Y" Then
    
        ' Restoring stock to a nil-stock lot
        If confirmNilStock Then
            resp = MsgBox("QC Number " & qcNumber & " is currently marked as Nil Stock." & vbCrLf & vbCrLf & _
                         "This transaction will restore the lot to active status. Do you want to continue?", _
                         vbYesNo + vbExclamation, "Restore from Nil Stock")
            If resp = vbNo Then
                result.ErrorMessage = "Transaction cancelled by user."
                GoTo CleanExit
            End If
        End If
        changeReasonText = piNumber & " " & reason & " - Restored from Nil Stock"
        updateResult = UpdateQCNumberRecord( _
            qcNumber:=qcNumber, _
            qcBalance:=newBalance, _
            nilStockFlag:="N", _
            nilStockFlagDate:="#CLEAR#", _
            nilStockFlagBy:="#CLEAR#", _
            changeReason:=changeReasonText, _
            qcTransID:=qcTransID)
    Else
        
        ' General update
                
        updateResult = UpdateQCNumberRecord( _
            qcNumber:=qcNumber, _
            qcBalance:=newBalance, _
            changeReason:=piNumber & " " & reason, _
            qcTransID:=qcTransID)
    End If

    If Not updateResult Then
        result.ErrorMessage = "Failed to update QC_Number.csv file."
        GoTo CleanExit
    End If

    ' --- STEP 12: Success ---
    result.Success = True
    GoTo CleanExit

ErrorHandler:
    result.Success = False
    result.ErrorMessage = "Error processing stock update: " & Err.Description
    
CleanExit:
    ProcessStockUpdate = result
End Function

' ========================================================================
' Helper: Get RM Code, Nil Stock Flag, and QC Status from QC Number
' ========================================================================
Private Function GetRMCodeFromQCNumber( _
    ByVal qcNumber As String, _
    ByRef rmCode As String, _
    ByRef nilStockFlag As String, _
    ByRef qcStatus As String) As Boolean
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim colQC As Long
    Dim colRM As Long
    Dim colNil As Long
    Dim colStatus As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("QC_Number")
    
    colQC = FindHeaderCol(ws, 1, HEADER_QC_NUMBER)
    colRM = FindHeaderCol(ws, 1, HEADER_RM_CODE)
    colNil = FindHeaderCol(ws, 1, HEADER_NIL_STOCK_FLAG)
    colStatus = FindHeaderCol(ws, 1, HEADER_QC_STATUS_FLAG)
    
    If colQC = 0 Or colRM = 0 Then
        GetRMCodeFromQCNumber = False
        GoTo CleanExit
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colQC).End(xlUp).row
    
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, colQC).value)) = qcNumber Then
            rmCode = Trim$(CStr(ws.Cells(r, colRM).value))
            
            If colNil > 0 Then
                nilStockFlag = Trim$(CStr(ws.Cells(r, colNil).value))
            Else
                nilStockFlag = "N"
            End If
            
            If colStatus > 0 Then
                qcStatus = Trim$(CStr(ws.Cells(r, colStatus).value))
            Else
                qcStatus = QC_STATUS_RELEASED
            End If
            
            GetRMCodeFromQCNumber = True
            GoTo CleanExit
        End If
    Next r
    
    GetRMCodeFromQCNumber = False
    GoTo CleanExit

ErrorHandler:
    GetRMCodeFromQCNumber = False

CleanExit:
    Set ws = Nothing
    Exit Function
End Function

' ========================================================================
' Helper: Get Master UoM and Conversion Factor
' ========================================================================
Private Function GetMasterUoMAndConversionFactor( _
    ByVal rmCode As String, _
    ByRef masterUoM As String, _
    ByRef wtPerML As Double) As Boolean
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim colRM As Long, colUoM As Long, colWt As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("RM_Master")
    
    colRM = FindHeaderCol(ws, 1, RMM_HEADER_RM_CODE)
    colUoM = FindHeaderCol(ws, 1, RMM_HEADER_MASTER_UOM)
    colWt = FindHeaderCol(ws, 1, RMM_HEADER_WT_PER_ML)
    
    If colRM = 0 Or colUoM = 0 Then
        GetMasterUoMAndConversionFactor = False
        GoTo CleanExit
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colRM).End(xlUp).row
    
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, colRM).value)) = rmCode Then
            masterUoM = Trim$(CStr(ws.Cells(r, colUoM).value))
            
            wtPerML = 0
            On Error Resume Next
            If colWt > 0 Then
                wtPerML = CDbl(ws.Cells(r, colWt).value)
            End If
            On Error GoTo ErrorHandler
            
            GetMasterUoMAndConversionFactor = True
            GoTo CleanExit
        End If
    Next r
    
    GetMasterUoMAndConversionFactor = False
    GoTo CleanExit

ErrorHandler:
    GetMasterUoMAndConversionFactor = False

CleanExit:
    Set ws = Nothing
    Exit Function
End Function

