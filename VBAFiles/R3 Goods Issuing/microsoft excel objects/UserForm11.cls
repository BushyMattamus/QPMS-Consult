Attribute VB_Name = "UserForm11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' ========================================================================
' GOODS ISSUE REVERSAL - WORKSHEET EVENTS
' ========================================================================
' Handles worksheet change events for the reversal form. Monitors cells
' D3 (Batch Number), D9 (Standard Name), and D11 (Inventory Transaction)
' to populate dropdowns, resolve RM Code, and load transaction details.
' ========================================================================

Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    
    On Error GoTo ErrorHandler
    
    ' --- Check if change is in Batch Number (D3) ---
    If Not Intersect(target, Me.Range("D3")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        ClearReversalDependentFields
        
        If Trim$(Me.Range("D3").value) <> "" Then
            DropdownUtils.PopulateStandardNameDropdownForBatch _
                Trim$(Me.Range("D3").value), Me.Range("D9")
            BusLogRecordLoader.LoadBatchDetailUsage
        Else
            DropdownUtils.ClearDropdownValidation Me.Range("D9")
        End If
        
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
        
    End If
    
    ' --- Check if change is in Standard Name (D9) ---
    If Not Intersect(target, Me.Range("D9")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        ClearReversalTransactionFields
        
        If Trim$(Me.Range("D9").value) <> "" Then
            BusLogRecordLoader.ResolveRMCodeFromDescription
            
            Dim batchNumber As String
            Dim rmCode As String
            
            batchNumber = Trim$(Me.Range("D3").value)
            rmCode = Trim$(Me.Range("D10").value)
            
            If rmCode <> "" Then
                BusLogReverseIssue.DisplayReversalTransactionReport batchNumber, rmCode
                DropdownUtils.PopulateTransactionDropdownForReversal batchNumber, rmCode, Me.Range("D11")
            End If
        Else
            Me.Range("D10").value = ""
            DropdownUtils.ClearDropdownValidation Me.Range("D11")
            ClearReversalTransactionReport
        End If
        
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
        
    End If
    
    ' --- Check if change is in Inventory Transaction (D11) ---
    If Not Intersect(target, Me.Range("D11")) Is Nothing Then
        
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        If Trim$(Me.Range("D11").value) <> "" Then
            BusLogRecordLoader.LoadReversalTransactionDetails Trim$(Me.Range("D11").value)
        Else
            ClearReversalDetailFields
        End If
        
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
        
    End If
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in Worksheet_Change: " & Err.Number & " - " & Err.Description, vbCritical, "UserForm11 Reversal"
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

CleanExit:
End Sub

'-------------------------------------------------------------------------------
' ClearReversalDependentFields
'-------------------------------------------------------------------------------
' Purpose: Clear all fields dependent on Batch Number selection
'-------------------------------------------------------------------------------
Private Sub ClearReversalDependentFields()
    Me.Range("D9").value = ""
    Me.Range("D10").value = ""
    Me.Range("D11").value = ""
    Me.Range("D12").value = ""
    Me.Range("D13").value = ""
    Me.Range("D14").value = ""
    Me.Range("D15").value = ""
    Me.Range("D16").value = ""
    
    On Error Resume Next
    Me.Range("D9").Validation.Delete
    Me.Range("D11").Validation.Delete
    On Error GoTo 0
    
    ClearReversalTransactionReport
End Sub

'-------------------------------------------------------------------------------
' ClearReversalTransactionFields
'-------------------------------------------------------------------------------
' Purpose: Clear fields dependent on Standard Name selection
'-------------------------------------------------------------------------------
Private Sub ClearReversalTransactionFields()
    Me.Range("D11").value = ""
    Me.Range("D12").value = ""
    Me.Range("D13").value = ""
    Me.Range("D14").value = ""
    Me.Range("D15").value = ""
    Me.Range("D16").value = ""
    
    On Error Resume Next
    Me.Range("D11").Validation.Delete
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' ClearReversalDetailFields
'-------------------------------------------------------------------------------
' Purpose: Clear detail fields (D12:D16) when transaction cleared
'-------------------------------------------------------------------------------
Private Sub ClearReversalDetailFields()
    Me.Range("D12").value = ""
    Me.Range("D13").value = ""
    Me.Range("D14").value = ""
    Me.Range("D15").value = ""
    Me.Range("D16").value = ""
End Sub

'-------------------------------------------------------------------------------
' ClearReversalTransactionReport
'-------------------------------------------------------------------------------
' Purpose: Clear the transaction report area (H11 onwards)
'-------------------------------------------------------------------------------
Private Sub ClearReversalTransactionReport()
    Dim lastRow As Long
    
    lastRow = Me.Cells(Me.Rows.Count, "H").End(xlUp).Row
    
    If lastRow >= 11 Then
        Me.Range("H11:L" & lastRow).ClearContents
    End If
End Sub
