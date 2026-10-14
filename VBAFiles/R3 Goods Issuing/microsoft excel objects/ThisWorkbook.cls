Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    
    ' Set workbook to read-only if in PROD environment
    If GetEnvironmentFromPath() = "PROD" Then
        ThisWorkbook.ChangeFileAccess xlReadOnly
    End If
    
    ' --- Performance Optimization ---
    AppStateUtils.DisableAppUpdates
    
    BusLogDataLoads.LoadQCNumberData
    BusLogDataLoads.FGBatchDetails
    BusLogDataLoads.LoadRMMasterData
    
    ' --- Re-enable events and screen updating ---
    AppStateUtils.RestoreAppState
        
End Sub

Private Sub Workbook_Activate()
    ' Enable drag and drop only in DEV mode, otherwise disable
    If PROD_DEV <> "DEV" Then
        Application.CellDragAndDrop = False
    End If
End Sub

Private Sub Workbook_Deactivate()
    ' Re-enable drag and drop when this workbook is no longer active (unless in DEV)
    Application.CellDragAndDrop = True
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error Resume Next
    
    ' Only perform backups if not in DEV mode
    If PROD_DEV <> "DEV" Then
        
        ' Only suppress save dialog if workbook is read-only
        If Me.ReadOnly Then
            Me.Saved = True
            Application.DisplayAlerts = False
        End If
        
        ' Backup frontend workbook (daily, overwrite)
        BackupFrontend
        
        ' Backup backend data tables (every close, timestamp, auto-cleanup)
        BackupDataTables
    End If
    
    On Error GoTo 0
End Sub
