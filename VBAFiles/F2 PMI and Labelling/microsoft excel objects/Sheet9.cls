Attribute VB_Name = "Sheet9"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Avery Maintenance
    ' This routine triggers when the Avery Code is changed, retrieves the corresponding Wide and High values from tblAvery,
    ' and populates them into the user form.

    ' Declare variables
    Dim wsLabels As Worksheet
    Dim tblAvery As ListObject
    Dim foundRow As ListRow
    Dim averyCode As String
    Dim wideValue As Long
    Dim highValue As Long

    ' Check if the change occurred in the Avery Code dropdown (D5)
    If Not Intersect(Target, Me.Range("D5")) Is Nothing Then
        ' Disable screen updating and events
        AppStateUtils.DisableAppUpdates
        
        On Error GoTo Cleanup

        ' Initialize variables
        Set wsLabels = ThisWorkbook.Worksheets("Avery Labels")
        Set tblAvery = wsLabels.ListObjects("tblAvery")
        averyCode = Me.Range("D5").value

        ' Locate the Avery Label Code in tblAvery
        Set foundRow = Nothing
        For Each foundRow In tblAvery.ListRows
            If foundRow.Range(1, 1).value = averyCode Then
                ' Retrieve the Wide and High values
                wideValue = foundRow.Range(1, 2).value
                highValue = foundRow.Range(1, 3).value
                Exit For
            End If
        Next foundRow

        ' Populate the Wide and High values in the form
        If Not foundRow Is Nothing Then
            Me.Range("D6").value = wideValue
            Me.Range("D7").value = highValue
        Else
            ' Clear the fields if no match is found
            Me.Range("D6").ClearContents
            Me.Range("D7").ClearContents
        End If

Cleanup:
        ' Re-enable screen updating and events
        AppStateUtils.RestoreAppState
    End If
End Sub
