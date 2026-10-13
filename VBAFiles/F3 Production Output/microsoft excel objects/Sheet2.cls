Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim batchNumber As String
    
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    ' --- Check if change was in C10 ---
    If Not Intersect(Target, Me.Range("C11")) Is Nothing Then
        ' --- Get batch number from form ---
        batchNumber = GetFieldValByLabelB("Batch Number:")
        
        ' --- Load batch details if batch number exists ---
        If Len(batchNumber) > 0 Then
            Call LoadBatchDetails(batchNumber)
        End If
    End If
    
    ' --- Success exit ---
    GoTo CleanExit
    
    ' --- Error handler ---
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Worksheet_Change"
    GoTo CleanExit
    
    ' --- Clean exit & object release ---
CleanExit:
    AppStateUtils.RestoreAppState
    FormUtils.ProtectActiveSheet
    Exit Sub
End Sub