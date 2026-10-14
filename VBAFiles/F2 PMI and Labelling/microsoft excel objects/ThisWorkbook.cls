Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    DisableAppUpdates
    On Error GoTo Cleanup

    ' Set workbook to read-only if in PROD environment
    If GetEnvironmentFromPath() = "PROD" Then
        ThisWorkbook.ChangeFileAccess xlReadOnly
    End If
    
    FGMasterDataLoad
    UpdateScannedBatchFlags
    FGBatchDetails
    LabelPrintsDataLoad
    LabelRegisterDataLoad

Cleanup:
    On Error Resume Next
    Application.StatusBar = False
    RestoreAppState
End Sub

Private Sub Workbook_Activate()
    ' Enable drag and drop only in DEV mode, otherwise disable
    If PROD_DEV <> "DEV" Then
        Application.CellDragAndDrop = False
    End If
End Sub

Private Sub Workbook_Deactivate()
    ' Re-enable drag and drop when this workbook is no longer active (unless in DEV)
    If PROD_DEV <> "DEV" Then
        Application.CellDragAndDrop = True
    End If
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

' ========================================================================
' UPDATE SCANNED BATCH FLAGS
' ========================================================================
' Orchestrates the scanned batch document verification process.
' Identifies unscanned batches, verifies PDF existence, and updates
' master and audit files accordingly.
' ========================================================================

Private Sub UpdateScannedBatchFlags()
    On Error GoTo ErrHandler

    Dim unscannedBatches() As String
    Dim foundBatches() As String
    Dim foundCount As Long
    Dim i As Long
    Dim updateSuccess As Boolean
    Dim successCount As Long
    Dim userName As String

    foundCount = 0
    successCount = 0
    userName = Environ("USERNAME")

    unscannedBatches = GetUnscannedBatches()

    If UBound(unscannedBatches) < 0 Then
        GoTo CleanExit
    End If

    foundBatches = CheckBatchPDFsExist(unscannedBatches)

    If UBound(foundBatches) < 0 Then
        GoTo CleanExit
    End If

    foundCount = UBound(foundBatches) + 1

    Application.StatusBar = "Updating Batch Doc Viewer. " & foundCount & " new file(s) found."

    For i = 0 To UBound(foundBatches)
        Application.StatusBar = "Updating batch " & (i + 1) & " of " & foundCount & ": " & foundBatches(i)
        
        updateSuccess = UpdateFinishedGoodsBatchRecord( _
            batchNumber:=foundBatches(i), _
            scanned:="x", _
            updatedBy:=userName, _
            updateComments:="Scanned batch document verified" _
        )
        
        If updateSuccess Then
            successCount = successCount + 1
        End If
    Next i

    If successCount > 0 Then
        Application.StatusBar = "Batch Doc Viewer updated. " & successCount & " batch(es) processed."
    End If

    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    Exit Sub
End Sub