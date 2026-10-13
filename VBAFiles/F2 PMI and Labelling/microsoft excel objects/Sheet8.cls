Attribute VB_Name = "Sheet8"
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

    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim qidCode As String
    Dim cpmuRef As String
    Dim userName As String
    Dim batchFrom As String
    Dim batchTo As String
    Dim mfgLotFrom As String
    Dim mfgLotTo As String
    Dim mfgLot2From As String
    Dim mfgLot2To As String

    If Intersect(Target, Me.Range("D5:D7, F5, I5:I8, K6:K8")) Is Nothing Then Exit Sub

    DisableAppUpdates

    dateFrom = ""
    dateTo = ""
    qidCode = Trim(Me.Range("D5").value)
    cpmuRef = Trim(Me.Range("D6").value)
    userName = Trim(Me.Range("G5").value)
    batchFrom = Trim(Me.Range("G6").value)
    batchTo = Trim(Me.Range("I6").value)
    mfgLotFrom = Trim(Me.Range("G7").value)
    mfgLotTo = Trim(Me.Range("I7").value)
    mfgLot2From = Trim(Me.Range("G8").value)
    mfgLot2To = Trim(Me.Range("I8").value)

    LoadPrintData qidCode, cpmuRef, dateFrom, dateTo, userName, batchFrom, batchTo, mfgLotFrom, mfgLotTo, mfgLot2From, mfgLot2To

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Worksheet_Change"
    GoTo CleanExit

CleanExit:
    RestoreAppState
    Exit Sub
End Sub