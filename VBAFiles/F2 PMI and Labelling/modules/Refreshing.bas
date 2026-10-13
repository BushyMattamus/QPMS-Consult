Attribute VB_Name = "Refreshing"
Sub PPRefresh()
    On Error GoTo ErrHandler
    
    Dim qidCode As String
    Dim cpmuRef As String
    Dim dateFrom As Variant
    Dim dateTo As Variant
    Dim userName As String
    Dim batchFrom As String
    Dim batchTo As String
    Dim mfgLotFrom As String
    Dim mfgLotTo As String
    Dim mfgLot2From As String
    Dim mfgLot2To As String
    Dim wsPrevious As Worksheet
    
    DisableAppUpdates

    LabelPrintsDataLoad
    
    Set wsPrevious = ThisWorkbook.Worksheets("Previous Prints")
    
    qidCode = Trim(wsPrevious.Range("D5").value)
    cpmuRef = Trim(wsPrevious.Range("D6").value)
    dateFrom = ""
    dateTo = ""
    userName = Trim(wsPrevious.Range("G5").value)
    batchFrom = Trim(wsPrevious.Range("G6").value)
    batchTo = Trim(wsPrevious.Range("I6").value)
    mfgLotFrom = Trim(wsPrevious.Range("G7").value)
    mfgLotTo = Trim(wsPrevious.Range("I7").value)
    mfgLot2From = Trim(wsPrevious.Range("G8").value)
    mfgLot2To = Trim(wsPrevious.Range("I8").value)
    
    LoadPrintData qidCode, cpmuRef, dateFrom, dateTo, _
        userName, batchFrom, batchTo, mfgLotFrom, mfgLotTo, _
        mfgLot2From, mfgLot2To
    
    RestoreAppState
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PPRefresh"
    GoTo CleanExit

CleanExit:
    Set wsPrevious = Nothing
    Exit Sub
End Sub

Sub NRefresh()
    On Error GoTo HandleError
    
    Dim qidCode As String
    Dim productID As String
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
    Dim wsPrints As Worksheet
    
    DisableAppUpdates

    LabelPrintsDataLoad
    
    Set wsPrints = ThisWorkbook.Worksheets("Reprint Labels")
    
    qidCode = Trim(wsPrints.Range("C5").value)
    cpmuRef = Trim(wsPrints.Range("C6").value)
    dateFrom = wsPrints.Range("F5").value
    dateTo = wsPrints.Range("H5").value
    userName = Trim(wsPrints.Range("F6").value)
    printedStatus = Trim(wsPrints.Range("F7").value)
    batchFrom = Trim(wsPrints.Range("K5").value)
    batchTo = Trim(wsPrints.Range("M5").value)
    mfgLotFrom = Trim(wsPrints.Range("K6").value)
    mfgLotTo = Trim(wsPrints.Range("M6").value)
    mfgLot2From = Trim(wsPrints.Range("K7").value)
    mfgLot2To = Trim(wsPrints.Range("M7").value)
    
    LoadPrintReprintData qidCode, cpmuRef, dateFrom, dateTo, _
        userName, printedStatus, batchFrom, batchTo, mfgLotFrom, mfgLotTo, _
        mfgLot2From, mfgLot2To
    
    RestoreAppState
    
    GoTo CleanExit

HandleError:
    MsgBox "Please close the dropdown or finish editing the cell before refreshing."
    GoTo CleanExit

CleanExit:
    Set wsPrints = Nothing
    Exit Sub
End Sub

Public Sub RefreshPMIViewer()
    On Error GoTo ErrHandler

    Dim qidCode As String
    Dim pmiNumber As String
    Dim batchNumber As String
    Dim productDescription As String
    Dim wsBatchViewer As Worksheet
    

    DisableAppUpdates

    FGBatchDetails

    Set wsBatchViewer = ThisWorkbook.Worksheets("Batch Doc Viewer")

    qidCode = Trim(wsBatchViewer.Range("D5").value)
    pmiNumber = Trim(wsBatchViewer.Range("D6").value)
    batchNumber = Trim(wsBatchViewer.Range("D7").value)
    productDescription = Trim(wsBatchViewer.Range("D8").value)

    LoadPMIViewerData qidCode, pmiNumber, batchNumber, productDescription

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "RefreshPMIViewer"
    GoTo CleanExit

CleanExit:
    RestoreAppState
    Set wsBatchViewer = Nothing
    Exit Sub
End Sub

