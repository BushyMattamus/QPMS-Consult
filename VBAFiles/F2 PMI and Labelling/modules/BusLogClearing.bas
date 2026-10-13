Attribute VB_Name = "BusLogClearing"
Sub AMEdClear()
    ' Avery Maintenance
    ' This sub disables events and screen updating, clears D5:D7, and then re-enables events and screen updating.
    
    On Error GoTo Cleanup

    ' Disable events and screen updating
    AppStateUtils.DisableAppUpdates

    ' Reference the Avery Maintenance worksheet and clear contents
    Worksheets("Avery Maintenance").Range("D5:D7").ClearContents

Cleanup:
    ' Re-enable events and screen updating
    AppStateUtils.RestoreAppState

    FormUtils.ProtectActiveSheet

End Sub

Sub AMNeClear()
    ' Avery Maintenance
    ' This sub disables events and screen updating, clears D16:D18, and then re-enables events and screen updating.
    
    On Error GoTo Cleanup

    ' Disable events and screen updating
    AppStateUtils.DisableAppUpdates

    ' Reference the Avery Maintenance worksheet and clear contents
    Worksheets("Avery Maintenance").Range("D16:D18").ClearContents

Cleanup:
    ' Re-enable events and screen updating
    AppStateUtils.RestoreAppState

    FormUtils.ProtectActiveSheet

End Sub

Sub BEClrNew()
    ' Batch Entry
    ' Clear / new button
    
    Dim ws As Worksheet

    ' Disable screen updating and events for performance
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    ' Set the target worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Assuming the button is on the same sheet as target cells

    ' Clear contents of specified cells
    ws.Range("D5:D7,D10:D12").ClearContents

    ' Re-enable screen updating and events
    AppStateUtils.RestoreAppState

    FormUtils.ProtectActiveSheet
    
End Sub

Sub BEClrPMIPRint()
    ' Label & PMI Print
    ' Clear Expiry Data
    
    Dim ws As Worksheet

    ' Disable screen updating and events for performance
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    ' Set the target worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Assuming the button is on the same sheet as target cells

    ' Clear contents of specified cells
    ws.Range("I5:I10, D5:D7, D10:D12").ClearContents

    ' Protect & Re-enable screen updating and events
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Sub BEClrBC()
    ' Batch Entry
    ' Clear Barcode
    
    Dim ws As Worksheet

    ' Disable screen updating and events for performance
    AppStateUtils.DisableAppUpdates

    ' Set the target worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Assuming the button is on the same sheet as target cells

    ' Clear contents of specified cells
    ws.Range("N5,N10").ClearContents

    ' Re-enable screen updating and events
    AppStateUtils.RestoreAppState

    FormUtils.ProtectActiveSheet
    
End Sub

Sub LMEdLabClr()
    ' Label Maintenance
    ' This sub disables events and screen updating, clears D5:D22, and then re-enables events and screen updating.
    
    Dim ws As Worksheet

    ' Disable events and screen updating
    AppStateUtils.DisableAppUpdates
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Reference the Label Maintenance worksheet and clear contents
    ws.Range("D5,D7:D9").ClearContents

Cleanup:
    ' Re-enable events and screen updating
    AppStateUtils.RestoreAppState

    ' Protect the worksheet
    FormUtils.ProtectActiveSheet

End Sub


Sub LMNwLabClr()
    ' Label Maintenance
    ' This VBA subroutine clears the contents of the range D34:D50 on the "Label Maintenance" worksheet.
    ' It disables events and screen updating during execution for performance and stability.

    Dim ws As Worksheet

    ' Disable events and screen updating for performance
    AppStateUtils.DisableAppUpdates
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Clear the specified range
    ws.Range("D20,D22:D24").ClearContents

Cleanup:
    ' Re-enable events and screen updating
    AppStateUtils.RestoreAppState

    ' Protect the worksheet
    FormUtils.ProtectActiveSheet


End Sub

Sub LRClear()
    ' Label Register Review
    ' This sub clears specific cell contents and resets the table on the "Label Register Review" worksheet.
    
    Dim wsReview As Worksheet
    Dim tblRegReview As ListObject
    Dim headerRange As Range
    Dim newTableRange As Range
    
    On Error GoTo Cleanup

    ' Disable events and screen updating
    AppStateUtils.DisableAppUpdates
    
    ' Unprotect the worksheet
    FormUtils.UnprotectActiveSheet
    
    ' Set the worksheet and table
    Set wsReview = ThisWorkbook.Worksheets("Label Register Review")
    Set tblRegReview = wsReview.ListObjects("tblRegReview")
    
    ' Define the range for the table's headers
    Set headerRange = wsReview.Range("B12:O12")
    
    ' Check if filters are applied to tblRegReview and clear them
    If tblRegReview.AutoFilter.FilterMode Then
        tblRegReview.AutoFilter.ShowAllData
    End If
    
    ' Clear the table's data
    On Error Resume Next
    If Not tblRegReview.DataBodyRange Is Nothing Then
        tblRegReview.DataBodyRange.ClearContents
    End If
    On Error GoTo 0
    
    ' Resize the table to just the header and one empty row
    Set newTableRange = headerRange.Resize(2) ' Add one empty row below the headers
    On Error Resume Next
    tblRegReview.Resize newTableRange
    On Error GoTo 0
    
    ' Clear additional specific ranges
    wsReview.Range("C5, C6, F5:F9, I5:I8, K6, K8").ClearContents

Cleanup:

    ' Protect the worksheet
    FormUtils.ProtectActiveSheet

    ' Re-enable events and screen updating
    AppStateUtils.RestoreAppState
End Sub

Sub PPClear()
    ' Previous Prints
    ' This sub clears specific cell contents and resets the table on the "Previous Prints" worksheet.
    
    Dim wsPreviousPrints As Worksheet
    Dim tblPreviousPrints As ListObject
    Dim headerRange As Range
    Dim newTableRange As Range
    
    On Error GoTo Cleanup

    ' Disable events and screen updating
    AppStateUtils.DisableAppUpdates
    
    ' Unprotect the worksheet
    FormUtils.UnprotectActiveSheet
    
    ' Set the worksheet and table
    Set wsPreviousPrints = ThisWorkbook.Worksheets("Previous Prints")
    Set tblPreviousPrints = wsPreviousPrints.ListObjects("tblPreviousPrints")
    
    ' Define the range for the table's headers
    Set headerRange = wsPreviousPrints.Range("B11:P11")
        
    ' Check if filters are applied to tblRegReview and clear them
    If tblPreviousPrints.AutoFilter.FilterMode Then
        tblPreviousPrints.AutoFilter.ShowAllData
    End If
    
    ' Clear the table's data
    On Error Resume Next
    If Not tblPreviousPrints.DataBodyRange Is Nothing Then
        tblPreviousPrints.DataBodyRange.ClearContents
    End If
    On Error GoTo 0
    
    ' Resize the table to just the header and one empty row
    Set newTableRange = headerRange.Resize(2) ' Add one empty row below the headers
    On Error Resume Next
    tblPreviousPrints.Resize newTableRange
    On Error GoTo 0
    
    ' Clear additional specific ranges
    wsPreviousPrints.Range("D5:D8, G5:G8, I6:I8").ClearContents

Cleanup:
    ' Protect the worksheet
    FormUtils.ProtectActiveSheet

    ' Re-enable events and screen updating
    AppStateUtils.RestoreAppState
End Sub

Sub PRClear()
    ' This sub clears specific cell contents and resets the table on the "Reprint Labels" worksheet.
    
    Dim wsPrints As Worksheet
    Dim tblPrints As ListObject
    Dim headerRange As Range
    Dim newTableRange As Range
    
    ' Set the worksheet and table
    Set wsPrints = ThisWorkbook.Worksheets("Reprint Labels")
    Set tblPrints = wsPrints.ListObjects("tblPrints")
        
    On Error Resume Next
    ActiveSheet.Range("C5").value = ActiveSheet.Range("C5").value
    If Err.Number <> 0 Then
        ' The user is likely in edit mode
        MsgBox "Please finish selecting from the dropdown or editing before clicking 'Clear'."
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo Cleanup
    
    ' Define the range for the table's headers
    Set headerRange = wsPrints.Range("B11:N11")
    
    ' Disable events and screen updating
    AppStateUtils.DisableAppUpdates
    
    ' Unprotect the worksheet
    FormUtils.UnprotectActiveSheet
    
    ' Check if filters are applied to tblRegReview and clear them
    If tblPrints.AutoFilter.FilterMode Then
        tblPrints.AutoFilter.ShowAllData
    End If
    
    ' Clear the table's data
    On Error Resume Next
    If Not tblPrints.DataBodyRange Is Nothing Then
        tblPrints.DataBodyRange.ClearContents
    End If
    On Error GoTo 0
    
    ' Resize the table to just the header and one empty row
    Set newTableRange = headerRange.Resize(2) ' Add one empty row below the headers
    On Error Resume Next
    tblPrints.Resize newTableRange
    On Error GoTo 0
    
    ' Clear additional specific ranges
    wsPrints.Range("C5:C7, F5:F7, H5, K5:K7, M5:M7").ClearContents

Cleanup:

    ' Protect the worksheet
    FormUtils.ProtectActiveSheet
    
    ' Re-enable events and screen updating
    AppStateUtils.RestoreAppState
End Sub

Sub PMIViewerClear()
    On Error GoTo ErrHandler
    
    Dim wsBatchViewer As Worksheet
    Dim tblPMIViewer As ListObject
    Dim headerRange As Range
    Dim newTableRange As Range
    
    Set wsBatchViewer = ThisWorkbook.Worksheets("Batch Doc Viewer")
    Set tblPMIViewer = wsBatchViewer.ListObjects("tblPMIViewer")
    
    Set headerRange = wsBatchViewer.Range("B11:H11")
    
    AppStateUtils.DisableAppUpdates
    
    wsBatchViewer.Unprotect
    
    If tblPMIViewer.AutoFilter.FilterMode Then
        tblPMIViewer.AutoFilter.ShowAllData
    End If
    
    On Error Resume Next
    If Not tblPMIViewer.DataBodyRange Is Nothing Then
        tblPMIViewer.DataBodyRange.ClearContents
    End If
    On Error GoTo ErrHandler
    
    Set newTableRange = headerRange.Resize(2)
    On Error Resume Next
    tblPMIViewer.Resize newTableRange
    On Error GoTo ErrHandler
    
    wsBatchViewer.Range("D5:D8").ClearContents
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PMIViewerClear"
    GoTo CleanExit

CleanExit:
    If PROD_DEV <> "DEV" Then
        wsBatchViewer.Protect _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            UserInterfaceOnly:=True
    End If
    
    AppStateUtils.RestoreAppState
    FormUtils.ProtectActiveSheet
    
    Set newTableRange = Nothing
    Set headerRange = Nothing
    Set tblPMIViewer = Nothing
    Set wsBatchViewer = Nothing
    Exit Sub
End Sub
