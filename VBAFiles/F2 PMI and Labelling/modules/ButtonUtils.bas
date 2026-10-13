Attribute VB_Name = "ButtonUtils"
Public Sub ButtSamplePMI()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogSampleView.SampleViewPMI

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Sample View PMI: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtViewPMI()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogTemplateViews.ViewPMI

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in View PMI: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtSavePrintPMI()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogSavePrint.SavePrintPMI
    Call BusLogDataLoads.FGBatchDetails

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Save Print PMI: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtRefreshBatchData()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogDataLoads.FGBatchDetails
    
    If Trim$(Range("I6").value) <> "" Then
        UnprotectActiveSheet
        Call BusLogRecordLoader.LoadBatchDetailsFromMaster
        ProtectActiveSheet
    End If
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Load Batch Details From Master: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtEditLabelSave()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogSaving.LMEdLabSave

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Edit Label Save: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtNewLabelSave()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogSaving.LMNwLabSave

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in New Label Save: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtDeleteLabelRegisterEntry()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogSaving.LMDelLabRec

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Delete Label Record: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtViewReprintLabelPDFs()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    ViewSelectedPDFs ARCHIVE_PRODUCT_LABELS_PATH, "Reprint Labels", "tblPrints", "Print ID", "View ('x')"

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in View Reprint Label PDFs: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtViewLabelPDFs()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    ViewSelectedPDFs ARCHIVE_PRODUCT_LABELS_PATH, "Previous Prints", "tblPreviousPrints", "Print ID", "View ('x')"

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in View Previous Print Label PDFs: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtViewOriginalBatch()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    ViewSelectedPDFs PRODUCTION_ARCHIVE_PMI_PATH, "Batch Doc Viewer", "tblPMIViewer", "Batch Number", "View ('x')"

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in View Previous Print Label PDFs: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtViewPMILabel()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    ViewSelectedPDFs ARCHIVE_PMI_LABELS_PATH, "Batch Doc Viewer", "tblPMIViewer", "Batch Number", "View ('x')"

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in View Previous Print Label PDFs: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub
Public Sub ButtViewCompletedBatch()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    ViewSelectedPDFs PRODUCTION_ARCHIVE_COMPLETED_PMI_PATH, "Batch Doc Viewer", "tblPMIViewer", "Batch Number", "View ('x')"

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in View Previous Print Label PDFs: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Public Sub ButtPrintPMIDocuments()
    AppStateUtils.DisableAppUpdates
    
    On Error GoTo ErrorHandler
    
    Call BusLogPrintPMI.PrintPMIDocuments

    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error in Print PMI Documents: " & Err.Description, vbCritical
    
CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub