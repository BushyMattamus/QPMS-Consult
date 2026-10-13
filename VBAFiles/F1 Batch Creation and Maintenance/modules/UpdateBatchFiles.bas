Attribute VB_Name = "UpdateBatchFiles"
' ========================================================================
' BATCH FILE UPDATES
' ========================================================================
' Manages updates to batch records in master and audit CSV files.
' Validates batch information, compares dates, and synchronizes changes
' across both files while maintaining audit trails.
' ========================================================================

Option Explicit

Public Sub UpdateBatchInCSV()
    On Error GoTo ErrHandler

    Const SHEET_NAME As String = "Batch Number Maintenance"
    Const BATCH_REF_SHEET As String = "BatchRef"

    Dim wsRef As Worksheet
    Dim batchNum As String
    Dim pmiNum As String
    Dim manDateText As String
    Dim expDateText As String
    Dim comments As String
    Dim batchRow As Long
    Dim newManDate As Date
    Dim newExpDate As Date
    Dim oldManDate As Date
    Dim oldExpDate As Date
    Dim updateSuccess As Boolean
    Dim newBatchStatus As String
    Dim oldBatchStatus As String
    Dim shelfNo As String
    Dim shelfLife As String
    Dim userResponse As VbMsgBoxResult

    Set wsRef = ThisWorkbook.Sheets(BATCH_REF_SHEET)

    ' --- Retrieve batch information ---
    batchNum = GetValue(SHEET_NAME, "Batch number")
    pmiNum = GetValue(SHEET_NAME, "PMI number")
    newBatchStatus = GetValue(SHEET_NAME, "Status")

    If batchNum = "" Then
        MsgBox "Batch Number cannot be empty.", vbExclamation, "UpdateBatchInCSV"
        GoTo CleanExit
    End If

    If pmiNum = "" Then
        MsgBox "PMI Number cannot be empty.", vbExclamation, "UpdateBatchInCSV"
        GoTo CleanExit
    End If

    ' --- Refresh selected CSV columns ---
    ImportFgBatchSelectedCSVColumns

    ' --- Find batch in reference sheet ---
    batchRow = GetBatchRow(wsRef, batchNum, pmiNum)
    If batchRow = 0 Then
        MsgBox "Batch not found in reference data.", vbCritical, "UpdateBatchInCSV"
        GoTo CleanExit
    End If

    ' --- Retrieve update fields ---
    manDateText = GetValue(SHEET_NAME, "Manufacturing date")
    expDateText = GetValue(SHEET_NAME, "Recommended expiry date")
    comments = GetValue(SHEET_NAME, "Comments")
    shelfNo = GetValue(SHEET_NAME, "Shelf life(No of days/Months/Years)")
    shelfLife = GetValue(SHEET_NAME, "Shelf life(Days/Months/Years)")

    ' --- Validate required fields ---
    If Not ValidateRequiredFields(manDateText, expDateText, comments) Then
        GoTo CleanExit
    End If

    If Not ValidateDates(manDateText, expDateText) Then
        GoTo CleanExit
    End If
    
    ' --- Comma validation for all fields being written ---
    If Not ValidateNoCommas(batchNum, "Batch Number") Then GoTo CleanExit
    If Not ValidateNoCommas(pmiNum, "PMI Number") Then GoTo CleanExit
    If Not ValidateNoCommas(GetValue(SHEET_NAME, "PMI version"), "PMI Version") Then GoTo CleanExit
    If Not ValidateNoCommas(GetValue(SHEET_NAME, "Product Description"), "Product Description") Then GoTo CleanExit
    If Not ValidateNoCommas(shelfNo, "Shelf life(No of days/Months/Years)") Then GoTo CleanExit
    If Not ValidateNoCommas(shelfLife, "Shelf life(Days/Months/Years)") Then GoTo CleanExit
    If Not ValidateNoCommas(comments, "Comments") Then GoTo CleanExit
    If Not ValidateNoCommas(newBatchStatus, "Status") Then GoTo CleanExit

    ' --- Validate expiry date calculation ---
    If Not ValidateExpiryDate(manDateText, expDateText, shelfNo, shelfLife, userResponse) Then
        If userResponse = vbNo Then
            GoTo CleanExit
        End If
    End If

    ' --- Convert and compare values ---
    newManDate = ConvertDDMMYYYY(manDateText)
    newExpDate = ConvertDDMMYYYY(expDateText)
    oldManDate = ConvertDDMMYYYY(ANSIToDate(wsRef.Cells(batchRow, 7).value))
    oldExpDate = ConvertDDMMYYYY(ANSIToDate(wsRef.Cells(batchRow, 8).value))
    oldBatchStatus = wsRef.Cells(batchRow, 11).value

    If newManDate = oldManDate And newExpDate = oldExpDate And newBatchStatus = oldBatchStatus Then
        MsgBox "No change detected for batch " & batchNum, vbInformation, "UpdateBatchInCSV"
        GoTo CleanExit
    End If

    ' --- Confirm critical status changes ---
    If newBatchStatus = FGB_STATUS_CANCELLED Or newBatchStatus = FGB_STATUS_REJECTED Then
        userResponse = MsgBox("You are about to change the batch status to '" & newBatchStatus & "'." & vbCrLf & vbCrLf & _
                              "This is a critical status change. Do you want to proceed?", _
                              vbYesNo + vbExclamation + vbDefaultButton2, "Confirm Status Change")
        
        If userResponse <> vbYes Then
            MsgBox "Batch update cancelled by user.", vbInformation, "UpdateBatchInCSV"
            GoTo CleanExit
        End If
    End If

    updateSuccess = UpdateFinishedGoodsBatchRecord( _
        batchNumber:=batchNum, _
        pmiNumber:=GetValue(SHEET_NAME, "PMI number"), _
        pmiVersion:=GetValue(SHEET_NAME, "PMI version"), _
        productDescription:=GetValue(SHEET_NAME, "Product Description"), _
        shelfLifeNo:=GetValue(SHEET_NAME, "Shelf life(No of days/Months/Years)"), _
        shelfLife:=GetValue(SHEET_NAME, "Shelf life(Days/Months/Years)"), _
        manufacturingDate:=DateToANSI(manDateText), _
        expiryDate:=DateToANSI(expDateText), _
        dateUpdated:=GetTransactionTimestamp(), _
        updatedBy:=Environ$("Username"), _
        updateComments:=comments, _
        batchStatus:=newBatchStatus _
    )

    If updateSuccess Then
        ImportFgBatchSelectedCSVColumns
        Helpers.ClearFields
        MsgBox "Batch record updated successfully.", vbInformation, "UpdateBatchInCSV"
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "UpdateBatchInCSV"
    GoTo CleanExit

CleanExit:
    Set wsRef = Nothing
End Sub

Private Function ValidateRequiredFields(manDate As String, expDate As String, comments As String) As Boolean
    On Error GoTo ErrHandler

    Dim missing As String

    If manDate = "" Then missing = missing & vbCrLf & "- Manufacturing Date"
    If expDate = "" Then missing = missing & vbCrLf & "- Recommended Expiry Date"
    If comments = "" Then missing = missing & vbCrLf & "- Comments"

    If missing <> "" Then
        MsgBox "Missing required fields:" & vbCrLf & missing, vbCritical, "ValidateRequiredFields"
        ValidateRequiredFields = False
        GoTo CleanExit
    End If

    ValidateRequiredFields = True
    GoTo CleanExit

ErrHandler:
    ValidateRequiredFields = False
    GoTo CleanExit

CleanExit:
End Function

Private Function ValidateDates(manDate As String, expDate As String) As Boolean
    On Error GoTo ErrHandler

    If Not IsValidDDMMYYYY(manDate) Or Not IsValidDDMMYYYY(expDate) Then
        MsgBox "Invalid date format (dd/mm/yyyy).", vbCritical, "ValidateDates"
        ValidateDates = False
        GoTo CleanExit
    End If

    If ConvertDDMMYYYY(expDate) <= ConvertDDMMYYYY(manDate) Then
        MsgBox "Expiry Date must be after Manufacturing Date.", vbCritical, "ValidateDates"
        ValidateDates = False
        GoTo CleanExit
    End If

    ValidateDates = True
    GoTo CleanExit

ErrHandler:
    ValidateDates = False
    GoTo CleanExit

CleanExit:
End Function

Public Function IsValidDDMMYYYY(dateText As String) As Boolean
    On Error Resume Next
    IsValidDDMMYYYY = IsDate(DateSerial(Split(dateText, "/")(2), _
                                        Split(dateText, "/")(1), _
                                        Split(dateText, "/")(0)))
End Function

Public Function ConvertDDMMYYYY(dateText As String) As Date
    ConvertDDMMYYYY = DateSerial(Split(dateText, "/")(2), _
                                 Split(dateText, "/")(1), _
                                 Split(dateText, "/")(0))
End Function

Private Function GetBatchRow(ws As Worksheet, batchNum As String, pmiNumber As String) As Long
    On Error GoTo ErrHandler

    Dim arr As Variant
    Dim i As Long

    arr = ws.Range("A1:B" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).value

    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = batchNum And arr(i, 2) = pmiNumber Then
            GetBatchRow = i
            GoTo CleanExit
        End If
    Next i

    GetBatchRow = 0
    GoTo CleanExit

ErrHandler:
    GetBatchRow = 0
    GoTo CleanExit

CleanExit:
End Function

Private Function GetValue(sheetName As String, label As String) As String
    GetValue = Trim(GetFieldByLabel(sheetName, "C", label).value)
End Function
