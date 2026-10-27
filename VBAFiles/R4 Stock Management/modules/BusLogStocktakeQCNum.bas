Attribute VB_Name = "BusLogStocktakeQCNum"
Option Explicit

' ========================================================================
' STOCKTAKE SINGLE QC NUMBER ENTRY
' ========================================================================
' Handles user input from "Single QC Number Stocktake" form
' Validates, processes, and posts the stocktake entry
' ========================================================================

Public Sub SubmitStocktakeQCNum()
    ' --- Variable declarations ---
    Dim standardName As String
    Dim rmCode As String
    Dim qcNumber As String
    Dim qcBalance As Double
    Dim masterUoM As String
    Dim countedQty As Double
    Dim countedUoM As String
    Dim difference As Double
    Dim notes As String
    Dim newRetestDate As String
    Dim countedBy As String
    Dim ok As Boolean
    Dim confirmMsg As String
    Dim nilStockFlagDate As String
    Dim lastCountDateLookup As Object
    Dim lastCountDate As String
    Dim tTotal As Double
    Dim tBlock As Double
    Dim tDialog As Double
    Dim dialogSec As Double
    
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    tTotal = PerfTimerNow()
    tBlock = tTotal
    dialogSec = 0
            
    ' --- STEP 1: Get all field values from form ---
    standardName = Trim$(GetFieldValByLabelC("Standard Name:"))
    rmCode = Trim$(GetFieldValByLabelC("RM Code:"))
    qcNumber = Trim$(GetFieldValByLabelC("QC Number:"))
    masterUoM = Trim$(GetFieldValByLabelC("Master UoM:"))
    countedUoM = Trim$(GetFieldValByLabelC("Count UoM:"))
    notes = Trim$(GetFieldValByLabelC("Notes:"))
    newRetestDate = Trim$(GetFieldValByLabelC("Actual Expiry/Retest Date:"))
    countedBy = Environ$("Username")
    PerfTimerLog "STOCKTAKE", "1 Read form fields", tBlock
    tBlock = PerfTimerNow()
    
    ' --- STEP 2: Validate mandatory fields ---
    If qcNumber = "" Then
        MsgBox "QC Number is required.", vbExclamation, "Missing Information"
        GoTo CleanExit
    End If
    
    If countedUoM = "" Then
        MsgBox "Count UoM is required.", vbExclamation, "Missing Information"
        GoTo CleanExit
    End If
    
    If GetFieldValByLabelC("Count Quantity:") = "" Then
        MsgBox "Count Quantity is required.", vbExclamation, "Missing Information"
        GoTo CleanExit
    End If
    
    If Not IsNumeric(GetFieldValByLabelC("Count Quantity:")) Then
        MsgBox "Count Quantity must be a valid number.", vbExclamation, "Invalid Input"
        GoTo CleanExit
    End If
    
    If Not ValidateNoCommas(notes, "Notes") Then GoTo CleanExit
    
    countedQty = CDbl(GetFieldValByLabelC("Count Quantity:"))
    PerfTimerLog "STOCKTAKE", "2 Validate inputs", tBlock
    tBlock = PerfTimerNow()
    
    ' --- STEP 3: Check nil stock exclusion logic ---
    nilStockFlagDate = GetQCNilStockFlagDate(qcNumber)
    PerfTimerLog "STOCKTAKE", "3a Nil stock flag lookup", tBlock
    tBlock = PerfTimerNow()
    
    If nilStockFlagDate <> "" Then
        Set lastCountDateLookup = GetLastCountDateLookup()
        PerfTimerLog "STOCKTAKE", "3b Last count date lookup", tBlock
        tBlock = PerfTimerNow()
        
        If lastCountDateLookup.Exists(qcNumber) Then
            lastCountDate = CStr(lastCountDateLookup(qcNumber))
            
            If DateToANSI(lastCountDate) >= DateToANSI(nilStockFlagDate) Then
                MsgBox "QC Number '" & qcNumber & "' has been flagged as nil stock and has already been counted since that flag was set." & vbCrLf & vbCrLf & _
                       "Nil Stock Flag Date: " & ANSIToDate(nilStockFlagDate) & vbCrLf & _
                       "Last Count Date: " & ANSIToDate(lastCountDate) & vbCrLf & vbCrLf & _
                       "This count cannot be posted.", _
                       vbExclamation, "Nil Stock Count Exists"
                GoTo CleanExit
            End If
        End If
    End If
    
    ' --- STEP 4: Check if already counted this month ---
    If IsQCCountedThisMonth(qcNumber) Then
        MsgBox "QC Number '" & qcNumber & "' has already been counted this month." & vbCrLf & vbCrLf & _
               "Only one count per QC Number per month is allowed.", _
               vbExclamation, "Duplicate Count"
        GoTo CleanExit
    End If
    PerfTimerLog "STOCKTAKE", "4 Monthly count check (incl. nil stock compare)", tBlock
    tBlock = PerfTimerNow()
        
    ' --- STEP 4: Get current balance from ADO connection ---
    qcBalance = GetQCBalanceADO(qcNumber)
    PerfTimerLog "STOCKTAKE", "4b Balance lookup (ADO)", tBlock
    tBlock = PerfTimerNow()
        
    ' Calculate difference
    difference = countedQty - qcBalance
    
    ' --- STEP 5: Show confirmation dialog ---
    confirmMsg = "Please confirm the stocktake entry:" & vbCrLf & vbCrLf & _
                 "QC Number: " & qcNumber & vbCrLf & _
                 "Standard Name: " & standardName & vbCrLf & _
                 "RM Code: " & rmCode & vbCrLf & vbCrLf & _
                 "Book Quantity: " & Format$(qcBalance, "0.000") & " " & masterUoM & vbCrLf & _
                 "Counted Quantity: " & Format$(countedQty, "0.000") & " " & countedUoM & vbCrLf & _
                 "Difference: " & Format$(difference, "0.000") & vbCrLf & vbCrLf & _
                 "Counted By: " & countedBy & vbCrLf & vbCrLf & _
                 "Do you want to post this count?"
    
    ' --- Confirmation dialog timing starts here ---
    tDialog = PerfTimerNow()
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Stocktake Entry") = vbNo Then
        dialogSec = dialogSec + (PerfTimerNow() - tDialog)
        GoTo CleanExit
    End If
        
    dialogSec = dialogSec + (PerfTimerNow() - tDialog)
        
    PerfTimerLog "STOCKTAKE", "5 Confirmation dialog (user wait)", tBlock
    tBlock = PerfTimerNow()

    ' --- STEP 6: Submit the stocktake entry ---
    ok = StocktakeQCNumEntry( _
        qcNumber:=qcNumber, _
        countedQty:=countedQty, _
        countedUoM:=countedUoM, _
        countedBy:=countedBy, _
        piNumber:="", _
        newRetestDate:=newRetestDate _
    )
        
    PerfTimerLog "STOCKTAKE", "6 StocktakeQCNumEntry write", tBlock
    tBlock = PerfTimerNow()
    
    If Not ok Then
        MsgBox "Failed to post stocktake entry. Please check the input and try again.", _
               vbCritical, "Stocktake Failed"
        GoTo CleanExit
    End If
    
    tDialog = PerfTimerNow()
    MsgBox "Stocktake entry posted successfully!" & vbCrLf & vbCrLf & _
           "QC Number: " & qcNumber & vbCrLf & _
           "Counted Quantity: " & Format$(countedQty, "0.000") & " " & countedUoM & vbCrLf & _
           "Book Quantity: " & Format$(qcBalance, "0.000") & " " & masterUoM & vbCrLf & _
           "Difference: " & Format$(difference, "0.000"), _
           vbInformation, "Stocktake Complete"
    dialogSec = dialogSec + (PerfTimerNow() - tDialog)
    PerfTimerLog "STOCKTAKE", "5b Success dialog (user wait)", tDialog
    
    ' Refresh data
    tBlock = PerfTimerNow()
    BusLogDataLoads.LoadQCNumberData

    PerfTimerLog "STOCKTAKE", "7 Refresh QC number data", tBlock
    tBlock = PerfTimerNow()

    ' Clear the form
    ClearStocktakeQC
    PerfTimerLog "STOCKTAKE", "8 Clear form", tBlock
    
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Submit Stocktake QC Num"
    GoTo CleanExit
    
CleanExit:
    PerfTimerLog "STOCKTAKE", "TOTAL (includes dialog waits)", tTotal
    Debug.Print "[STOCKTAKE TIMER] CODE-ONLY TOTAL (excludes dialog waits): " & _
                Format$(((PerfTimerNow() - tTotal) - dialogSec) * 1000#, "0.000") & " ms"
    Set lastCountDateLookup = Nothing
    Exit Sub
End Sub

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

