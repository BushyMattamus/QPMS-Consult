Attribute VB_Name = "Sheet11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler

    Dim qidCode As String
    Dim pmiNumber As String
    Dim batchNumber As String
    Dim productDescription As String

    If Intersect(Target, Me.Range("D5:D8")) Is Nothing Then Exit Sub

    DisableAppUpdates

    qidCode = Trim(Me.Range("D5").value)
    pmiNumber = Trim(Me.Range("D6").value)
    batchNumber = Trim(Me.Range("D7").value)
    productDescription = Trim(Me.Range("D8").value)

    LoadPMIViewerData qidCode, pmiNumber, batchNumber, productDescription

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Worksheet_Change"
    GoTo CleanExit

CleanExit:
    RestoreAppState
    Exit Sub
End Sub