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
    On Error GoTo ErrorHandler
    
    ' Set workbook to read-only if in PROD environment
    If GetEnvironmentFromPath() = "PROD" Then
        ThisWorkbook.ChangeFileAccess xlReadOnly
    End If
    
    ' --- Performance Optimization ---
    DisableAppUpdates
    
    Call LoadQCNumberData
    
    ' --- Re-enable events and screen updating ---
    RestoreAppState
    
Exit Sub

ErrorHandler:
    MsgBox "Error in Workbook_Activate: " & Err.Description, vbCritical, "Error"
    RestoreAppState
    
End Sub

Private Sub Workbook_Activate()
    On Error GoTo ErrorHandler

    ' Enable drag and drop only in DEV mode, otherwise disable
    If PROD_DEV <> "DEV" Then
        Application.CellDragAndDrop = False
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error in Workbook_Activate: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Workbook_Deactivate()
    On Error GoTo ErrorHandler

    ' Re-enable drag and drop when this workbook is no longer active
    Application.CellDragAndDrop = True

    Exit Sub

ErrorHandler:
    MsgBox "Error in Workbook_Deactivate: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrorHandler

    ' Only perform backups if not in DEV mode
    If PROD_DEV <> "DEV" Then
        
        ' Only suppress save dialog if workbook is read-only
        If Me.ReadOnly Then
            Me.Saved = True
            Application.DisplayAlerts = False
        End If
        
        ' Backup frontend workbook (daily, overwrite)
        BackupUtils.BackupFrontend
        
        ' Backup backend data tables (every close, timestamp, auto-cleanup)
        BackupUtils.BackupDataTables
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error in Workbook_BeforeClose: " & Err.Description, vbCritical, "Error"
End Sub
