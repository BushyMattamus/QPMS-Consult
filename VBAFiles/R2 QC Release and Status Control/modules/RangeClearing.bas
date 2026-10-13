Attribute VB_Name = "RangeClearing"
Public Sub ClearMaintainQCNumber()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D21").ClearContents

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Sub ClearExport()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D4:D13").ClearContents

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Sub ClearLiveReport()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D15").ClearContents
    
    GenerateLiveQCReport

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

    ActiveSheet.Range("D3").value = ActiveSheet.Range("D3").value
    
End Sub

Public Sub ClearQCRelease()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D14").ClearContents

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub