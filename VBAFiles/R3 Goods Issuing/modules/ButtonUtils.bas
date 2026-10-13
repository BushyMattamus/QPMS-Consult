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

Public Sub ButtLoadBatchDetailsData()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.FGBatchDetails
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load Batch Details Data: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtSaveGoodsIssue()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogIssueGoods.SaveGoodsIssue
    
    ' --- Reload & Cleanup ---
    BusLogDataLoads.LoadQCNumberData
    BusLogDataLoads.FGBatchDetails
    
    RangeClearing.ClearIssuingRM
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Save Goods Issue: " & Err.Description, vbCritical

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateWhereUsedReport()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogExport.GenerateWhereUsedReport

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Where Used Report: " & Err.Description, vbCritical

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtGenerateBatchMaterialsReport()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogBatchMaterial.GenerateBatchMaterialsReport

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Generate Batch Materials Report: " & Err.Description, vbCritical

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ButtSaveGoodsReversal()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogReverseIssue.SaveGoodsReversal

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Save Goods Reversal: " & Err.Description, vbCritical

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub