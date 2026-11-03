Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)

    ' Label Maintenance

    Dim ws As Worksheet
    Dim wsMaintenance As Worksheet
    Dim tbl As ListObject
    Dim pmiNumber As String
    Dim foundRow As ListRow
    
    ' Only run if the change is in cell D5
    If Intersect(Target, Me.Range("D5")) Is Nothing Then Exit Sub

    ' Ensure cleanup will be performed at the end of this sub
    On Error GoTo Cleanup

    ' Exit if the change is not in D5
    If Not Intersect(Target, Me.Range("D5")) Is Nothing Then

        ' Disable events and screen updating for performance
        DisableAppUpdates
        
        UnprotectActiveSheet
        
        ' Get the PMI Number from D5
        lookupPMINumber = Target.value

        ' Reference the worksheet and table
        Set ws = Worksheets("Label Register")
        Set tbl = ws.ListObjects("tblLabelReg")
        Set wsMaintenance = Worksheets("Label Maintenance")

        ' Clear existing return fields
        wsMaintenance.Range("D7").ClearContents

        ' Check if PMI Number exists in the table
        For Each foundRow In tbl.ListRows
            If foundRow.Range(1).value = lookupPMINumber Then ' Assuming PMI Numberis in the 1st column
                ' Populate the fields in the return fields section
                wsMaintenance.Range("D7").value = foundRow.Range(2).value ' Avery Label Code
                wsMaintenance.Range("D8").value = foundRow.Range(6).value ' Avery Label Code
                wsMaintenance.Range("D9").value = foundRow.Range(5).value ' Avery Label Code
                Exit For
            End If
        Next foundRow

    End If

Cleanup:

    ProtectActiveSheet

    RestoreAppState
    
End Sub