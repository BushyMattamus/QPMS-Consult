Attribute VB_Name = "UserForm30"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Run LoadQCNumberRecord when cell D3 changes
    If Not Intersect(Target, Me.Range("D3")) Is Nothing Then
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        
        Range("D4:D18").ClearContents
        
        LoadQCNumberRecord
        AppStateUtils.RestoreAppState
        FormUtils.ProtectActiveSheet
    End If
End Sub