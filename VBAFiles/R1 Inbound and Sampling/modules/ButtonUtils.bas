Attribute VB_Name = "ButtonUtils"
Option Explicit

Public Sub ButtRefreshRMMaster()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnlockWorkbook
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.LoadRMMasterData
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load RM Master Data: " & Err.Description, vbCritical
    
CleanExit:
    AppStateUtils.RestoreAppState
    FormUtils.LockWorkbook
End Sub

Public Sub ButtRefreshLiveReport()
    DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.LoadQCNumberData
    
    ActiveSheet.Range("D3").value = ActiveSheet.Range("D3").value
    
    RestoreAppState
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load QC Number Data: " & Err.Description, vbCritical
    
CleanExit:
End Sub

Public Sub ButtRefreshQCNumber()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    FormUtils.UnlockWorkbook
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.LoadQCNumberData
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load QC Number Data: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    FormUtils.LockWorkbook
    AppStateUtils.RestoreAppState
End Sub

Public Sub ButtRefreshVendorList()
    DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.LoadVendorList
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load Vendor List: " & Err.Description, vbCritical
    
CleanExit:
    RestoreAppState
End Sub

Public Sub ButtRefreshManufacturerList()
    DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.LoadManufacturerList
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load Manufacturer List: " & Err.Description, vbCritical
    
CleanExit:
    RestoreAppState
End Sub

Public Sub ButtExportLiveReport()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    Call BusLogExport.ExportLiveReport
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Export Live Report: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
End Sub

Public Sub ButtExportQCNumbers()

    Dim wsDataExport As Worksheet

    On Error GoTo ErrorHandler

    Set wsDataExport = ThisWorkbook.Sheets("Data Export")

    DisableAppUpdates
    wsDataExport.Unprotect
    
    Call BusLogExport.ExportQCNumbers
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Export QC Numbers: " & Err.Description, vbCritical
    
CleanExit:
    If PROD_DEV <> "DEV" Then
        wsDataExport.Protect
    End If

    RestoreAppState
End Sub

Public Sub ButtExportQCNumberChanges()

    Dim wsDataExport As Worksheet

    On Error GoTo ErrorHandler

    Set wsDataExport = ThisWorkbook.Sheets("Data Export")

    DisableAppUpdates
    wsDataExport.Unprotect
    
    Call BusLogExport.ExportQCNumberChanges
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Export QC Number Changes: " & Err.Description, vbCritical
    
CleanExit:
    If PROD_DEV <> "DEV" Then
        wsDataExport.Protect
    End If

    RestoreAppState
End Sub

Public Sub ButtSaveMaintainQCNumber()
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    FormUtils.UnlockWorkbook
    
    On Error GoTo ErrorHandler
    
    Call BusLogMaintainQCNumber.SaveMaintainQCNumber
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Save Maintain QC Number: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.LockWorkbook
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
End Sub

Public Sub ButtRefreshRequiresSampling()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    Call BusLogRequiresSampling.RequiresSampling
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Requires Sampling Refresh: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
End Sub

Public Sub ButtSaveGR()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    Call BusLogSaving.SaveGR
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Saving Goods Receipt: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
End Sub

Public Sub ButtSaveQCSample()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    Call BusLogSaving.SaveQCSample
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Saving Sampling: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
End Sub

Public Sub ButtSaveVendor()
    DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogSaving.SaveVendorList
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Saving Vendor List: " & Err.Description, vbCritical
    
CleanExit:
    RestoreAppState
End Sub

Public Sub ButtSaveManufacturer()
    DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogSaving.SaveManufacturerList
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Saving Manufacturer List: " & Err.Description, vbCritical
    
CleanExit:
    RestoreAppState
End Sub

Public Sub ButtNewQCNumber()
    DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    UnprotectActiveSheet
        
    Call BusLogInboundCreation.GetNextQCNumber
    
    ProtectActiveSheet
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in NewQCNumber: " & Err.Description, vbCritical
    
CleanExit:
    RestoreAppState
End Sub