Attribute VB_Name = "ButtonUtils"
Public Sub ButtRefreshQCNumberData()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.LoadQCNumberData

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load QC Number Data: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtSaveStockAdjustment()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogAdjustments.SaveStockAdjustment

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Save Stock Adjustment: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtSubmitStocktakeQCNum()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    

    Call BusLogStocktakeQCNum.SubmitStocktakeQCNum

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Submit Stocktake QC Number: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateStocktakeCountList()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogStocktakeMassCount.GenerateStocktakeCountList

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Stocktake Count List: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtSubmitStocktakeMassCountSheet()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogStocktakeMassCount.SubmitStocktakeMassCountSheet
    Call BusLogDataLoads.LoadQCNumberData

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Submit Stocktake Mass Count Sheet: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateMonthlyWriteOffsReport()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogReportingMonthsWriteOffs.GenerateMonthlyWriteOffsReport

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Monthly Write Offs Report: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateRetestAnalysisReport()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogReportingExpiryProfile.GenerateRetestAnalysisReport

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Retest Analysis Report: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateManufacturerExpiryReport()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogReportingExpiryProfile.GenerateManufacturerExpiryReport

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Manufacturer Expiry Report: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateBalancesReport()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogReportingBalances.GenerateBalancesReport

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Balances Report: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateCurrentSOHReport()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogReportingBalances.GenerateCurrentSOHReport

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Current SOH Report: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub