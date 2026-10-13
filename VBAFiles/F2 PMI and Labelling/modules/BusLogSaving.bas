Attribute VB_Name = "BusLogSaving"
Sub AMEdSave()
    ' Avery Maintenance
    ' This routine saves the user changes to tblAvery, overwriting the relevant row or adding a new one.

    ' Declare variables
    Dim wsLabels As Worksheet
    Dim tblAvery As ListObject
    Dim foundRow As ListRow
    Dim averyCode As String
    Dim wideValue As Long
    Dim highValue As Long
    Dim totalValue As Long
    Dim userName As String
    Dim currentDate As Date
    Dim wsMaintenance As Worksheet
    Dim pivotTable As pivotTable

    ' Disable screen updating and events
    AppStateUtils.DisableAppUpdates

    On Error GoTo Cleanup

    ' Initialize variables
    Set wsLabels = ThisWorkbook.Worksheets("Avery Labels")
    Set tblAvery = wsLabels.ListObjects("tblAvery")
    Set wsMaintenance = ThisWorkbook.Worksheets("Avery Maintenance")
    averyCode = wsMaintenance.Range("D5").Text
    wideValue = wsMaintenance.Range("D6").value
    highValue = wsMaintenance.Range("D7").value
    totalValue = wideValue * highValue
    userName = Environ("Username")
    currentDate = Now

    ' Locate the Avery Label Code in tblAvery
    Set foundRow = Nothing
    For Each foundRow In tblAvery.ListRows
        If foundRow.Range(1, 1).value = averyCode Then
            ' Overwrite the existing row
            foundRow.Range(1, 2).value = wideValue
            foundRow.Range(1, 3).value = highValue
            foundRow.Range(1, 4).value = totalValue
            foundRow.Range(1, 5).value = userName
            foundRow.Range(1, 6).value = currentDate
            GoTo RefreshPivot
        End If
    Next foundRow

RefreshPivot:
    
    ' Unprotect the sheet
    FormUtils.UnprotectActiveSheet
    
    ' Refresh the pivot table in Avery Maintenance
    For Each pivotTable In wsMaintenance.PivotTables
        pivotTable.RefreshTable
    Next pivotTable

    ' Clear the contents of cells D5:D7
    wsMaintenance.Range("D5:D7").ClearContents

Cleanup:
    
    ' Protect the worksheet
    If PROD_DEV <> "DEV" Then
        FormUtils.ProtectActiveSheet
    End If

    ' Re-enable screen updating and events
    AppStateUtils.RestoreAppState
End Sub

Sub AMNeSave()
    ' Avery Maintenance
    ' This routine saves new label details entered by the user to tblAvery, ensuring the Avery Code is unique.

    ' Declare variables
    Dim wsLabels As Worksheet
    Dim tblAvery As ListObject
    Dim foundRow As ListRow
    Dim averyCode As String
    Dim wideValue As Long
    Dim highValue As Long
    Dim totalValue As Long
    Dim userName As String
    Dim currentDate As Date
    Dim wsMaintenance As Worksheet
    Dim pivotTable As pivotTable

    ' Disable screen updating and events
    AppStateUtils.DisableAppUpdates

    On Error GoTo Cleanup

    ' Initialize variables
    Set wsLabels = ThisWorkbook.Worksheets("Avery Labels")
    Set tblAvery = wsLabels.ListObjects("tblAvery")
    Set wsMaintenance = ThisWorkbook.Worksheets("Avery Maintenance")
    averyCode = wsMaintenance.Range("D16").Text
    wideValue = wsMaintenance.Range("D17").value
    highValue = wsMaintenance.Range("D18").value
    totalValue = wideValue * highValue
    userName = Environ("Username")
    currentDate = Now

    ' Validate that the Avery Code is not already in use
    For Each foundRow In tblAvery.ListRows
        If foundRow.Range(1, 1).value = averyCode Then
            MsgBox "The Avery Code " & averyCode & " is already in use. Please enter a unique code.", vbExclamation, "Duplicate Avery Code"
            GoTo Cleanup
        End If
    Next foundRow

    ' Add a new row to tblAvery
    Set foundRow = tblAvery.ListRows.Add
    foundRow.Range(1, 1).value = averyCode
    foundRow.Range(1, 2).value = wideValue
    foundRow.Range(1, 3).value = highValue
    foundRow.Range(1, 4).value = totalValue
    foundRow.Range(1, 5).value = userName
    foundRow.Range(1, 6).value = currentDate
    
    ' Unprotect the sheet
    FormUtils.UnprotectActiveSheet
    
    ' Refresh the pivot table in Avery Maintenance
    For Each pivotTable In wsMaintenance.PivotTables
        pivotTable.RefreshTable
    Next pivotTable

    ' Clear the contents of D16:D18
    wsMaintenance.Range("D16:D18").ClearContents

Cleanup:
    
    ' Protect the worksheet
    If PROD_DEV <> "DEV" Then
        FormUtils.ProtectActiveSheet
    End If
    
    ' Re-enable screen updating and events
    AppStateUtils.RestoreAppState
End Sub

Sub LMEdLabSave()

    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Dim pmiNumber As String
    Dim averyLabelCode As String
    Dim specialInstructions As String
    Dim labelsPerBatch As String
    Dim inputValid As Boolean
    Dim updateSuccess As Boolean

    AppStateUtils.DisableAppUpdates

    Set wsSource = Worksheets("Label Maintenance")

    pmiNumber = Trim$(GetFieldValByLabelC("PMI Number:"))
    averyLabelCode = Trim$(GetFieldValByLabelC("Avery Label Code:"))
    specialInstructions = Trim$(GetFieldValByLabelC("Special Instructions:"))
    labelsPerBatch = Trim$(GetFieldValByLabelC("Labels Per Batch:"))

    inputValid = True
    
    If pmiNumber = "" Then
        MsgBox "Error: PMI Number cannot be empty.", vbExclamation, "Validation Error"
        inputValid = False
        GoTo CleanExit
    End If
    
    ' Validate Labels Per Batch if provided
    If labelsPerBatch <> "" Then
        If Not IsNumeric(labelsPerBatch) Then
            MsgBox "Error: Labels Per Batch must be a valid number.", vbExclamation, "Validation Error"
            inputValid = False
            GoTo CleanExit
        End If
    End If

    If Not inputValid Then GoTo CleanExit

    updateSuccess = UpdateLabelRegisterRecord( _
        pmiNumber:=pmiNumber, _
        averyCode:=averyLabelCode, _
        specialInstructions:=specialInstructions, _
        labelsPerBatch:=labelsPerBatch)

    If updateSuccess Then
        MsgBox "Label details saved successfully.", vbInformation, "Success"
        LabelRegisterDataLoad
        LMEdLabClr
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LMEdLabSave"
    GoTo CleanExit

CleanExit:
    Set wsSource = Nothing
    AppStateUtils.RestoreAppState
    Exit Sub
End Sub

Sub LMNwLabSave()

    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Dim pmiNumber As String
    Dim averyLabelCode As String
    Dim specialInstructions As String
    Dim labelsPerBatch As String
    Dim inputValid As Boolean
    Dim saveSuccess As Boolean

    AppStateUtils.DisableAppUpdates

    Set wsSource = Worksheets("Label Maintenance")

    pmiNumber = Trim$(GetFieldValByLabelC("New PMI Number:"))
    averyLabelCode = Trim$(GetFieldValByLabelC("New Avery Label Code:"))
    specialInstructions = Trim$(GetFieldValByLabelC("New Special Instructions:"))
    labelsPerBatch = Trim$(GetFieldValByLabelC("New Labels Per Batch:"))

    inputValid = True
    
    If pmiNumber = "" Then
        MsgBox "Error: PMI Number cannot be empty.", vbExclamation, "Validation Error"
        inputValid = False
        GoTo CleanExit
    End If
    
    If averyLabelCode = "" Then
        MsgBox "Error: Avery Label Code cannot be empty.", vbExclamation, "Validation Error"
        inputValid = False
        GoTo CleanExit
    End If
    
    If labelsPerBatch <> "" Then
        If Not IsNumeric(labelsPerBatch) Then
            MsgBox "Error: Labels Per Batch must be a valid number.", vbExclamation, "Validation Error"
            inputValid = False
            GoTo CleanExit
        End If
    End If

    If Not inputValid Then GoTo CleanExit

    saveSuccess = AppendLabelRegisterRecord( _
        pmiNumber:=pmiNumber, _
        averyCode:=averyLabelCode, _
        specialInstructions:=specialInstructions, _
        labelsPerBatch:=labelsPerBatch)

    If saveSuccess Then
        MsgBox "New label details saved successfully.", vbInformation, "Success"
        LabelRegisterDataLoad
        LMNwLabClr
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LMNwLabSave"
    GoTo CleanExit

CleanExit:
    Set wsSource = Nothing
    AppStateUtils.RestoreAppState
    Exit Sub
End Sub

Sub LMDelLabRec()

    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Dim pmiNumber As String
    Dim deleteSuccess As Boolean
    Dim confirmResult As VbMsgBoxResult

    AppStateUtils.DisableAppUpdates

    Set wsSource = Worksheets("Label Maintenance")

    pmiNumber = Trim$(GetFieldValByLabelC("PMI Number:"))

    If pmiNumber = "" Then
        MsgBox "Error: PMI Number cannot be empty.", vbExclamation, "Validation Error"
        GoTo CleanExit
    End If

    confirmResult = MsgBox("Are you sure you want to delete the label record for PMI Number: " & pmiNumber & "?", _
                           vbYesNo + vbQuestion, "Confirm Deletion")

    If confirmResult = vbNo Then
        GoTo CleanExit
    End If

    deleteSuccess = UpdateLabelRegisterRecord( _
        pmiNumber:=pmiNumber, _
        deleteRecord:=True)

    If deleteSuccess Then
        MsgBox "Label record deleted successfully.", vbInformation, "Success"
        LabelRegisterDataLoad
        LMEdLabClr
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LMDelLabRec"
    GoTo CleanExit

CleanExit:
    Set wsSource = Nothing
    AppStateUtils.RestoreAppState
    Exit Sub
End Sub
