Attribute VB_Name = "UserForm20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)

    ' Only respond to changes in D6
    If Intersect(Target, Me.Range("D6")) Is Nothing Then Exit Sub

    On Error GoTo CleanUp

    AppStateUtils.DisableAppUpdates
    FormUtils.UnprotectActiveSheet

    BusLogDataLoads.LoadQCNumberData
    Call BusLogRecordLoader.LoadRMDetailsFromQC(Me.Range("D6"), Me.Range("D4"), Me.Range("D5"))

    GetQCBalanceADO Me.Range("D6")

    ' Ensure formulas have recalculated before reading D8
    Me.Calculate

    ' Replicate calculated value
    Me.Range("D10").value = Me.Range("D8").value

CleanUp:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState

End Sub