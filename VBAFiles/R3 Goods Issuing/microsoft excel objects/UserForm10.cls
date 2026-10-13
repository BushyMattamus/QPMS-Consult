Attribute VB_Name = "UserForm10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal target As Range)
    
    ' Check if the change is in cell D3 (Batch Number), return batch details
    If Not Intersect(target, Me.Range("D3")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        If Range("D3") = "" Then
            Range("D4:D6").ClearContents
        Else
            BusLogRecordLoader.LoadBatchDetailUsage
        End If
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
    End If
    
    ' Check if the change is in cell D9 (RM Description), return unique RM code
    If Not Intersect(target, Me.Range("D9")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        If Range("D9") = "" Then
            Range("D10").ClearContents
        Else
            BusLogRecordLoader.ResolveRMCodeFromDescription
        End If
        
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
        
        ActiveSheet.Range("D10").value = ActiveSheet.Range("D10").value
        
    End If
    
    ' Check if the change is in cell D10 (RM Code), return a list of QC numbers
    If Not Intersect(target, Me.Range("D10")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        BusLogDataLoads.LoadQCNumberData
        BusLogRecordLoader.PopulateQCNumberList
        
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
    End If
    
    ' Check if the change is in cell D11 (QC Number), return QC number details
    If Not Intersect(target, Me.Range("D11")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        BusLogDataLoads.LoadQCNumberData
        BusLogRecordLoader.LoadRMDetailsFromQC
        
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
    End If
    
End Sub