Attribute VB_Name = "ButtonUtils"
Sub BTNClearFields()

DisableAppUpdates
UnprotectActiveSheet

Helpers.ClearFields
 
ProtectActiveSheet
RestoreAppState

End Sub

Sub BTNCreateBatch()

DisableAppUpdates
UnprotectActiveSheet

WriteBatchToCSV ("Batch Number Generation")

ProtectActiveSheet
RestoreAppState
End Sub

Sub BTNLoadBatch()

DisableAppUpdates
UnprotectActiveSheet

LoadBatchForMaintenance

ProtectActiveSheet
RestoreAppState
End Sub

Sub BTNloadproduct()

DisableAppUpdates
UnprotectActiveSheet

LoadProductFields ("Batch Number Generation")

ProtectActiveSheet
RestoreAppState

End Sub

Sub BTNUpdateBatch()

DisableAppUpdates
UnprotectActiveSheet

UpdateBatchInCSV
    
ProtectActiveSheet
RestoreAppState

End Sub

Sub BTNRefreshData()

DisableAppUpdates
UnprotectActiveSheet

LoadFgMaster.ImportSelectedCSVColumns
LoadFgBatch.ImportFgBatchSelectedCSVColumns
    
ProtectActiveSheet
RestoreAppState

End Sub