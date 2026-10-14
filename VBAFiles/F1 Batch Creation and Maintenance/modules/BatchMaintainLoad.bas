Attribute VB_Name = "BatchMaintainLoad"
Public Sub LoadBatchForMaintenance()

    Dim wsRef As Worksheet
    Dim sheetName As String
    Dim batchNum As String
    Dim rowNum As Variant

    On Error GoTo ErrHandler

    Set wsRef = ThisWorkbook.Sheets("BatchRef")

    sheetName = "Batch Number Maintenance"

    ' --- Read batch number from screen ---
    batchNum = Trim(GetFieldByLabel(sheetName, "C", "Batch number").value)

    If batchNum = "" Then
        ClearProductFields
        Exit Sub
    End If

    ' --- Find the batch number in BatchRef (column 1) ---
    rowNum = Application.Match(batchNum, wsRef.Columns(1), 0)

    If IsError(rowNum) Then
        MsgBox "Batch Number '" & batchNum & "' not found.", vbCritical
        Exit Sub
    End If

    ' --- Populate the values on the Batch Number Maintenance sheet ---
    With wsRef

        GetFieldByLabel(sheetName, "C", "PMI number").value = .Cells(rowNum, 2).value
        GetFieldByLabel(sheetName, "C", "PMI version").value = .Cells(rowNum, 3).value
        GetFieldByLabel(sheetName, "C", "Product description").value = .Cells(rowNum, 4).value
        GetFieldByLabel(sheetName, "C", "Shelf life(No of days/Months/Years)").value = .Cells(rowNum, 5).value
        GetFieldByLabel(sheetName, "C", "Shelf life(Days/Months/Years)").value = .Cells(rowNum, 6).value
        GetFieldByLabel(sheetName, "C", "Manufacturing date").value = .Cells(rowNum, 7).value
        GetFieldByLabel(sheetName, "C", "Recommended expiry date").value = .Cells(rowNum, 8).value
        GetFieldByLabel(sheetName, "C", "Status").value = .Cells(rowNum, 11).value
        
    End With

   ' MsgBox "Batch details loaded successfully.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in LoadBatchForMaintenance: " & Err.Description, vbCritical

End Sub

