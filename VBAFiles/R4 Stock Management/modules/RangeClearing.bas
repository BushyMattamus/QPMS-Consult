Attribute VB_Name = "RangeClearing"
Public Sub ClearStockAdjust()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D4:D7, D11:D13, D15").ClearContents
    
    Range("D4").Select

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ClearStocktakeQC()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D4:D6, D9:D10, D13:D14").ClearContents
    
    Range("D6").Select

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub


Public Sub ClearStocktakeSheet()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D4:D5").ClearContents
    BusLogDataLoads.LoadQCNumberData

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub