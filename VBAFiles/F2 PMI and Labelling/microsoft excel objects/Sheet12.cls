Attribute VB_Name = "Sheet12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim qidCode As String
    Dim cpmuRef As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim userName As String
    Dim printedStatus As String
    Dim batchFrom As String
    Dim batchTo As String
    Dim mfgLotFrom As String
    Dim mfgLotTo As String
    Dim mfgLot2From As String
    Dim mfgLot2To As String

    If Intersect(Target, Me.Range("C5:M7")) Is Nothing Then Exit Sub

    DisableAppUpdates
    UnprotectActiveSheet

    qidCode = Trim(Me.Range("C5").value)
    cpmuRef = Trim(Me.Range("C6").value)
    dateFrom = Me.Range("F5").value
    dateTo = Me.Range("H5").value
    userName = Trim(Me.Range("F6").value)
    printedStatus = Trim(Me.Range("F7").value)
    batchFrom = Trim(Me.Range("K5").value)
    batchTo = Trim(Me.Range("M5").value)
    mfgLotFrom = Trim(Me.Range("K6").value)
    mfgLotTo = Trim(Me.Range("M6").value)
    mfgLot2From = Trim(Me.Range("K7").value)
    mfgLot2To = Trim(Me.Range("M7").value)

    LoadPrintReprintData qidCode, cpmuRef, dateFrom, dateTo, _
        userName, printedStatus, batchFrom, batchTo, mfgLotFrom, mfgLotTo, _
        mfgLot2From, mfgLot2To

    ProtectActiveSheet
    RestoreAppState
End Sub
