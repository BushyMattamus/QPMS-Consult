Attribute VB_Name = "RangeClearing"
Public Sub ClearIssuingRM()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D9:D11,D16:D20,D22").ClearContents
    Range("D18").value = "No"

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Sub ClearIssuingAll()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D11,D16:D20,D22").ClearContents
    Range("D18").value = "No"

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
End Sub

Sub ClearLiveReport()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D15").ClearContents
    
    GenerateLiveQCReport

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

    ActiveSheet.Range("D3").value = ActiveSheet.Range("D3").value
    
End Sub

Public Sub ClearReversalAll()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D3:D17").ClearContents

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ClearWhereUsed()

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    Range("D4:D7").ClearContents

    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub

Public Sub ClearReversalFormFields()
    Dim batchNumber As String
    
    On Error Resume Next
    
    ' --- Get batch number before clearing ---
    batchNumber = Trim$(ActiveSheet.Range("D3").value)
    On Error Resume Next
    
    AppStateUtils.DisableAppUpdates
    UnprotectActiveSheet
    
    With ActiveSheet
        .Range("D9:D17").value = ""
    End With
    
    ClearDropdownValidation ActiveSheet.Range("D9")
    ClearDropdownValidation ActiveSheet.Range("D11")

    ' --- Repopulate Standard Name dropdown if batch still selected ---
    If batchNumber <> "" Then
        PopulateStandardNameDropdownForBatch batchNumber, ActiveSheet.Range("D9")
    End If
    
    ProtectActiveSheet
    AppStateUtils.RestoreAppState
    
    On Error GoTo 0
End Sub