Attribute VB_Name = "ButtonUtils"
Public Sub ButtRefreshQCNumberData()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    Call LoadQCNumberData
    Call GenerateLiveQCReport
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load QC Number Data: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
    
End Sub

Public Sub ButtExportLiveReport()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    BusLogExport.ExportLiveReport
    
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
Public Sub ButtSaveQCStatus()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    BusLogMaintainQCStatus.SaveQCStatus
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Save QC Status: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
    
End Sub

Public Sub ButtSaveQCRelease()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    BusLogMaintainQCStatus.SaveQCRelease
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Save QC Release: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
    
End Sub

Public Sub ButtNewTesting()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    BusLogRequiresTesting.NewTesting
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in New Testing: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
    
End Sub

Public Sub ButtUpcomingRetests()
    DisableAppUpdates
    UnprotectActiveSheet
    
    On Error GoTo ErrorHandler
    
    BusLogRequiresTesting.UpcomingRetests
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Upcoming Retests: " & Err.Description, vbCritical
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
    
End Sub