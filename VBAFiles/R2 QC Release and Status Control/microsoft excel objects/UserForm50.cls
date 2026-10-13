Attribute VB_Name = "UserForm50"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal target As Range)
    ' Run GenerateLiveQCReport when cell D3 changes
    If Not Intersect(target, Me.Range("D3, D4, D6:D8, D10:D12, D14, D15")) Is Nothing Then
        DisableAppUpdates
        UnprotectActiveSheet
        
        Call GenerateLiveQCReport
        
        ProtectActiveSheet
        RestoreAppState
    End If
End Sub