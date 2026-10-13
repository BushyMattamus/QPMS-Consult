Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_PivotTableUpdate(ByVal target As PivotTable)
    Dim pt As PivotTable
    Dim ptRange As Range
    Dim printRange As Range
    
    On Error GoTo ErrHandler
    
    If target.Name = "PivotTable1" Then
        Set pt = target
        Set ptRange = pt.TableRange2
        
        Set printRange = Range(ptRange.Cells(-2, 1), _
                               ptRange.Cells(ptRange.Rows.Count, ptRange.Columns.Count + 1))
        
        Me.PageSetup.PrintArea = printRange.Address
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Worksheet_PivotTableUpdate"
    GoTo CleanExit

CleanExit:
    Set printRange = Nothing
    Set ptRange = Nothing
    Set pt = Nothing
End Sub
