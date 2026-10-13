Attribute VB_Name = "UserForm20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if the changed cell is D3
    If Not Intersect(Target, Me.Range("D3")) Is Nothing Then
        ' Prevent infinite loop if the macro changes D3
        AppStateUtils.DisableAppUpdates
        FormUtils.UnprotectActiveSheet
        On Error GoTo CleanExit
        
        Range("D7:D9").ClearContents

        Dim qcNumber As String
        Dim balance As Double

        qcNumber = Me.Range("D3").value
        If qcNumber <> "" Then
            balance = GetQCBalanceFromRMTrans(qcNumber)
            Me.Range("D6").value = balance
        Else
            Me.Range("D6").value = ""
        End If

CleanExit:
        FormUtils.ProtectActiveSheet
        AppStateUtils.RestoreAppState
    End If
End Sub