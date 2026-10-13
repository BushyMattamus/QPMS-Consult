Attribute VB_Name = "BusLogDeleting"
Sub AMEdDel()
    ' Avery Maintenance
    ' This routine deletes an existing Avery label entry from tblAvery using the fields in the Edit Avery Template section.

    ' Declare variables
    Dim wsLabels As Worksheet
    Dim tblAvery As ListObject
    Dim foundRow As ListRow
    Dim averyCode As String
    Dim wsMaintenance As Worksheet
    Dim pivotTable As pivotTable

    ' Disable screen updating and events
    AppStateUtils.DisableAppUpdates

    On Error GoTo Cleanup

    ' Initialize variables
    Set wsLabels = ThisWorkbook.Worksheets("Avery Labels")
    Set tblAvery = wsLabels.ListObjects("tblAvery")
    Set wsMaintenance = ThisWorkbook.Worksheets("Avery Maintenance")
    averyCode = wsMaintenance.Range("D5").value
    
    ' Unprotect the sheet
    FormUtils.UnprotectActiveSheet

    ' Validate that the Avery Code is provided
    If averyCode = "" Then
        MsgBox "Please select an Avery Code to delete.", vbExclamation, "Missing Avery Code"
        GoTo Cleanup
    End If

    ' Locate the Avery Label Code in tblAvery
    Set foundRow = Nothing
    For Each foundRow In tblAvery.ListRows
        If foundRow.Range(1, 1).value = averyCode Then
            ' Prompt the user for confirmation
            If MsgBox("Are you sure you want to delete the Avery Label Code " & averyCode & "?", vbOKCancel + vbExclamation, "Confirm Delete") = vbOK Then
                ' Delete the row
                foundRow.Delete
                MsgBox "Avery Label Code " & averyCode & " has been deleted.", vbInformation, "Deleted Successfully"
            End If
            GoTo RefreshPivot
        End If
    Next foundRow

    ' If no match is found
    MsgBox "The Avery Code " & averyCode & " was not found in tblAvery.", vbExclamation, "Not Found"

RefreshPivot:
    ' Refresh the pivot table in Avery Maintenance
    For Each pivotTable In wsMaintenance.PivotTables
        pivotTable.RefreshTable
    Next pivotTable

    ' Clear the contents of D5:D7
    wsMaintenance.Range("D5:D7").ClearContents

Cleanup:
    
    ' Reprotect the sheet
    FormUtils.ProtectActiveSheet

    ' Re-enable screen updating and events
    AppStateUtils.RestoreAppState
End Sub