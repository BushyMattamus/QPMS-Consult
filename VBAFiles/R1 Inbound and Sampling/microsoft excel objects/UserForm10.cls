Attribute VB_Name = "UserForm10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Clear RM code on Standard Name change
    If Not Intersect(Target, Me.Range("D5")) Is Nothing Then
        Range("D6").ClearContents
    End If
    
    If Not Intersect(Target, Me.Range("D4")) Is Nothing Then
        ' Get dynamic values from the form
        
        DisableAppUpdates
        UnprotectActiveSheet
        
        ' Update search results - output starts destination cell
        Application.EnableEvents = False
        UpdateFuzzySearchResults Target, Me.Range("L3"), 20
        Application.EnableEvents = True
        
        ProtectActiveSheet
        RestoreAppState
        
    End If
    
    ' Check if the change is in cell D5 (RM Description), return unique RM code
    If Not Intersect(Target, Me.Range("D5")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        If Range("D5") = "" Then
            Range("D6").ClearContents
        Else
            Call BusLogRecordLoader.ResolveRMCodeFromDescription(Range("D5"), Range("D6"))
        End If
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
        
        ActiveSheet.Range("D6").value = ActiveSheet.Range("D6").value
        
    End If
    
End Sub