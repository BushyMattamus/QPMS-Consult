Attribute VB_Name = "BusLogAdjustments"
Public Sub SaveStockAdjustment()
    Dim qcNumber As String
    Dim adjustQtyStr As String
    Dim adjUoM As String
    Dim adjReason As String
    Dim adjType As String
    Dim adjustQty As Double
    Dim result As StockUpdateResult
    
    On Error GoTo ErrorHandler

    BusLogDataLoads.LoadQCNumberData
    
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
    
    ' Call centralized stock update function
    result = ProcessStockUpdate( _
        qcNumber:=qcNumber, _
        inputQty:=adjustQty, _
        inputUoM:=adjUoM, _
        reason:=adjReason, _
        adjustmentType:=adjType, _
        confirmNilStock:=True _
    )
    
    If Not result.Success Then
        MsgBox result.ErrorMessage, vbExclamation, "Stock Adjustment Failed"
        Exit Sub
    End If
    
    ' Reload and cleanup
    BusLogDataLoads.LoadQCNumberData
    RangeClearing.ClearStockAdjust
    
    ' Show confirmation
    Dim actionDesc As String
    actionDesc = IIf(result.adjustmentType = "WRITE-OFF", "written off", "written on")
    
    MsgBox "Raw Material has been " & actionDesc & "." & vbCrLf & vbCrLf & _
           "QC Number: " & qcNumber & vbCrLf & _
           "Quantity: " & Format$(Abs(result.transactionQty), "0.000") & vbCrLf & _
           "Previous Balance: " & Format$(result.PreviousBalance, "0.000") & vbCrLf & _
           "New Balance: " & Format$(result.newBalance, "0.000"), _
           vbInformation, "Adjustment Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving stock adjustment: " & Err.Description, vbCritical
    
CleanExit:
    ' No object variables to release in this procedure.
    Exit Sub
End Sub
