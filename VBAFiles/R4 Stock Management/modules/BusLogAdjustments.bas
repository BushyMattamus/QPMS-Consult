Attribute VB_Name = "BusLogAdjustments"
Public Sub SaveStockAdjustment()
    Dim qcNumber As String
    Dim adjustQtyStr As String
    Dim adjUoM As String
    Dim adjReason As String
    Dim adjType As String
    Dim adjustQty As Double
    Dim result As StockUpdateResult
    Dim tTotal As Double
    Dim tBlock As Double
    
    On Error GoTo ErrorHandler
    tTotal = PerfTimerNow()

    BusLogDataLoads.LoadQCNumberData
    
    PerfTimerLog "ADJUST", "Load QC number data", tTotal
    tBlock = PerfTimerNow()

    ' Get Fields
    qcNumber = Trim$(GetFieldValByLabelC("QC Number:"))
    adjustQtyStr = Trim$(GetFieldValByLabelC("Adjustment Quantity:"))
    adjUoM = Trim$(GetFieldValByLabelC("Adjustment UoM:"))
    adjReason = Trim$(GetFieldValByLabelC("Adjustment Reason:"))
    adjType = Trim$(GetFieldValByLabelC("Adjustment Type:"))
    
    ' --- Validate no Commas ---
    If Not ValidateNoCommas(adjReason, "Adjustment Reason") Then GoTo CleanExit
    
    ' Validate mandatory fields
    If qcNumber = "" Or adjustQtyStr = "" Or adjUoM = "" Or adjType = "" Or adjReason = "" Then
        MsgBox "Please complete all mandatory fields.", vbExclamation, "Missing Information"
        GoTo CleanExit
    End If
    
    If Not IsNumeric(adjustQtyStr) Then
        MsgBox "Adjustment Quantity must be a valid number.", vbExclamation
        GoTo CleanExit
    End If
    
    adjustQty = CDbl(adjustQtyStr)
    
    PerfTimerLog "ADJUST", "Read fields, validate, parse quantity", tBlock
    tBlock = PerfTimerNow()

    ' Call centralized stock update function
    result = ProcessStockUpdate( _
        qcNumber:=qcNumber, _
        inputQty:=adjustQty, _
        inputUoM:=adjUoM, _
        reason:=adjReason, _
        adjustmentType:=adjType, _
        confirmNilStock:=True _
    )
    
    PerfTimerLog "ADJUST", "ProcessStockUpdate", tBlock
    tBlock = PerfTimerNow()

    If Not result.Success Then
        PerfTimerLog "ADJUST", "TOTAL (failed, incl. dialog)", tTotal
                MsgBox result.ErrorMessage, vbExclamation, "Stock Adjustment Failed"
        Exit Sub
    End If
    
    ' Reload and cleanup
    BusLogDataLoads.LoadQCNumberData
    RangeClearing.ClearStockAdjust
    PerfTimerLog "ADJUST", "Reload and clear form", tBlock
    tBlock = PerfTimerNow()
    
    ' Show confirmation (suppressed in test mode; see IsTestMode in PathResolver)
    PerfTimerLog "ADJUST", "TOTAL (excl. completion dialog)", tTotal
    Dim actionDesc As String
    actionDesc = IIf(result.adjustmentType = "WRITE-OFF", "written off", "written on")
    
    Dim confirmMsg As String
    confirmMsg = "Raw Material has been " & actionDesc & "." & vbCrLf & vbCrLf & _
           "QC Number: " & qcNumber & vbCrLf & _
           "Quantity: " & Format$(Abs(result.transactionQty), "0.000") & vbCrLf & _
           "Previous Balance: " & Format$(result.PreviousBalance, "0.000") & vbCrLf & _
           "New Balance: " & Format$(result.newBalance, "0.000")
    
    If IsTestMode() Then
        Debug.Print "Adjustment Complete (test mode): " & Replace$(confirmMsg, vbCrLf, " ")
    Else
        MsgBox confirmMsg, vbInformation, "Adjustment Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving stock adjustment: " & Err.Description, vbCritical
    
CleanExit:
    PerfTimerLog "ADJUST", "TOTAL (error path, incl. dialog)", tTotal
    ' No object variables to release in this procedure.
    Exit Sub
End Sub
