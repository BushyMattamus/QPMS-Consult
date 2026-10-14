Attribute VB_Name = "BusLogClearing"
Option Explicit

Sub Clear_ProdOutput_Form()
    '-------------------------------------
    ' Purpose: Clears all data entry fields on the input form.
    '-------------------------------------
    
    ' Declare worksheet variable
    Dim wsInput As Worksheet

    ' Set to your input sheet name
    Set wsInput = ThisWorkbook.Sheets("Production Entry")
    
    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet
    
    ' Clear each input cell as specified
    wsInput.Range("C5:C13, C17:C18, C23:E29").ClearContents
    wsInput.Range("C19").value = ""
    
    Set wsInput = Nothing
    
    AppStateUtils.RestoreAppState
    FormUtils.ProtectActiveSheet
    
End Sub
