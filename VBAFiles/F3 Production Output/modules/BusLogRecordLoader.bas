Attribute VB_Name = "BusLogRecordLoader"
' ========================================================================
' BATCH DETAILS LOADER
' ========================================================================
' Loads batch details from the Batch_Details worksheet
' Retrieves Quick_Code, Product_Description, and PMI_number based on
' the provided Batch_Number and populates target cells
' ========================================================================

Option Explicit

Public Sub LoadBatchDetails(ByVal batchNumber As String)
    ' --- Error handling setup ---
    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim wsBatchDetails As Worksheet
    Dim wsActive As Worksheet
    Dim batchRow As Long
    Dim quickCode As String
    Dim productDescription As String
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim batchStatus As String

    ' --- Object initialization ---
    Set wsActive = ActiveSheet
    Set wsBatchDetails = ThisWorkbook.Sheets("Batch_Details")

    ' --- Validate batch number ---
    If Len(batchNumber) = 0 Then
        MsgBox "Batch Number is required.", vbExclamation, "LoadBatchDetails"
        GoTo CleanExit
    End If

    ' --- Find batch row in Batch_Details ---
    batchRow = 0
    On Error Resume Next
    batchRow = Application.Match(batchNumber, wsBatchDetails.Columns(1), 0)
    On Error GoTo ErrHandler

    If batchRow = 0 Then
        MsgBox "Batch Number '" & batchNumber & "' not found in Batch_Details.", vbExclamation, "LoadBatchDetails"
        GoTo CleanExit
    End If

    ' --- Extract batch details ---
    quickCode = wsBatchDetails.Cells(batchRow, 16).value
    productDescription = wsBatchDetails.Cells(batchRow, 4).value
    pmiNumber = wsBatchDetails.Cells(batchRow, 2).value
    pmiVersion = wsBatchDetails.Cells(batchRow, 3).value
    batchStatus = wsBatchDetails.Cells(batchRow, 11).value


    ' --- Populate target cells ---
    wsActive.Range("C5").value = quickCode
    wsActive.Range("C6").value = productDescription
    wsActive.Range("C7").value = pmiNumber
    wsActive.Range("C8").value = pmiVersion
    wsActive.Range("C9").value = batchStatus

    ' --- Success exit ---
    GoTo CleanExit

    ' --- Error handler ---
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LoadBatchDetails"
    GoTo CleanExit

    ' --- Clean exit & object release ---
CleanExit:
    Set wsBatchDetails = Nothing
    Set wsActive = Nothing
    Exit Sub
End Sub
