Attribute VB_Name = "UserForm10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Check if the change is in cell D9 (RM Description), return unique RM code
    If Not Intersect(Target, Me.Range("D4")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        If Range("D4") = "" Then
            Range("D5").ClearContents
        Else
            Call BusLogRecordLoader.ResolveRMCodeFromDescription(Range("D4"), Range("D5"))
        End If
        
        Call BusLogRecordLoader.PopulateQCNumberList(Range("D5"), Range("H5:M100000"))
        
        AppStateUtils.RestoreAppState
        FormUtils.ProtectActiveSheet
        
        Range("D6").Select
        
    End If
    
    ' Check if the change is in cell D5 (RM Code), return a list of QC numbers
    If Not Intersect(Target, Me.Range("D5")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        Call BusLogRecordLoader.PopulateQCNumberList(Range("D5"), Range("H5:M100000"))
        
        AppStateUtils.RestoreAppState
        FormUtils.ProtectActiveSheet
        
    End If
    
    ' Check if the change is in cell D6 (QC Number), return QC number details
    If Not Intersect(Target, Me.Range("D6")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        UnprotectActiveSheet
        
        BusLogDataLoads.LoadQCNumberData
        Call BusLogRecordLoader.LoadRMDetailsFromQC(Range("D6"), Range("D4"), Range("D5"))
        Me.Range("D7").value = GetQCBalanceADO(Range("D6").value)
        
        ' Ensure formulas have recalculated before reading D8
        Me.Calculate

        ' Replicate calculated value
        Me.Range("D13").value = Me.Range("D8").value
        
        ProtectActiveSheet
        AppStateUtils.RestoreAppState
        
    End If
    
End Sub