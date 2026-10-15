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
    Dim stageStartedAt As Double

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    On Error GoTo ErrorHandler

    Call BusLogIssueGoods.SaveGoodsIssue
    
    ' --- Reload & Cleanup ---
    stageStartedAt = Timer
    Debug.Print "[R3 Goods Issue] START LoadQCNumberData at " & Format$(Now, "hh:nn:ss")
    BusLogDataLoads.LoadQCNumberData
    Debug.Print "[R3 Goods Issue] END LoadQCNumberData; elapsed " & Format$(ElapsedSeconds(stageStartedAt), "0.000") & " seconds"

    stageStartedAt = Timer
    Debug.Print "[R3 Goods Issue] START FGBatchDetails at " & Format$(Now, "hh:nn:ss")
    BusLogDataLoads.FGBatchDetails
    Debug.Print "[R3 Goods Issue] END FGBatchDetails; elapsed " & Format$(ElapsedSeconds(stageStartedAt), "0.000") & " seconds"
    
    stageStartedAt = Timer
    Debug.Print "[R3 Goods Issue] START ClearIssuingRM at " & Format$(Now, "hh:nn:ss")
    RangeClearing.ClearIssuingRM
    Debug.Print "[R3 Goods Issue] END ClearIssuingRM; elapsed " & Format$(ElapsedSeconds(stageStartedAt), "0.000") & " seconds"
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Save Goods Issue: " & Err.Description, vbCritical

CleanExit:
    stageStartedAt = Timer
    Debug.Print "[R3 Goods Issue] START final protection and app-state restore at " & Format$(Now, "hh:nn:ss")
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    Debug.Print "[R3 Goods Issue] END final protection and app-state restore; elapsed " & Format$(ElapsedSeconds(stageStartedAt), "0.000") & " seconds"

End Sub

Private Function ElapsedSeconds(ByVal startedAt As Double) As Double
    ElapsedSeconds = Timer - startedAt
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function

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