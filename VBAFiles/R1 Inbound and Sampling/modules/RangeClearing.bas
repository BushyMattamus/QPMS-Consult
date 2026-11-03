Attribute VB_Name = "RangeClearing"
Public Sub ClearSample()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3, D6:D9").ClearContents
    
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub


Public Sub ClearGR()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D4:D6, D10:D16, D18:D23").ClearContents
    
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub


Public Sub ClearMaintainQCNumber()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D20").ClearContents
    
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Sub ClearExport()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D4:D14").ClearContents
    
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Sub ClearLiveReport()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D15").ClearContents
    
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

    ActiveSheet.Range("D3").value = ActiveSheet.Range("D3").value
    
End Sub