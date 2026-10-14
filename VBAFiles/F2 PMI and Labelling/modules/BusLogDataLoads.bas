Attribute VB_Name = "BusLogDataLoads"
' ========================================================================
' BUSINESS LOGIC DATA LOADS
' ========================================================================
' Handles importing CSV data files into worksheet tables for:
' - Finished Goods Batch Details
' - Finished Goods Master Data
' - Label Prints Data
' - Print/Reprint filtering and display
' ========================================================================

Option Explicit

' ========================================================================
' FG Batch Details Import with Date Conversion
' ========================================================================
' Imports the Finished Goods Batch Master file into the "Batch_Details" sheet.
' Converts date columns (Manufacturing_Date, Expiry_Date, Date_Created) to dd/mm/yyyy format.
' ========================================================================

Public Sub FGBatchDetails()
    On Error GoTo ErrHandler

    Dim fgbdCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long

    DisableAppUpdates

    ' --- Resolve source CSV path and destination sheet ---
    fgbdCSVPath = GetWorkbookTableFilePath(FILE_FINISHED_GOODS_BATCH_MASTER)
    Set wsDest = ThisWorkbook.Worksheets("Batch_Details")

    ' --- Clear existing contents and import CSV ---
    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(fgbdCSVPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 105, , ERROR_CSV_IMPORT & fgbdCSVPath
    End If

    ' --- Update named range NRbatchNum to A2:A(lastRow) ---
    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    If lastRow >= 2 Then
        ThisWorkbook.Names("NRbatchNum").RefersTo = wsDest.Range("A2:A" & lastRow)
        ThisWorkbook.Names("NRbatchPMINum").RefersTo = wsDest.Range("B2:B" & lastRow)
        ThisWorkbook.Names("NRbatchProdDesc").RefersTo = wsDest.Range("D2:D" & lastRow)
        ThisWorkbook.Names("NRbatchStatus").RefersTo = wsDest.Range("K2:K" & lastRow)
        
        ' --- Populate Quick_Code in column P using FG_Master lookup ---
        wsDest.Cells(1, 17).value = "Quick_Code"
        PopulateQuickCodeFromMaster wsDest, lastRow
        
        ThisWorkbook.Names("NRbatchQuickCode").RefersTo = wsDest.Range("Q2:Q" & lastRow)
        
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "FG Batch Detail Load Error: " & Err.Description, vbCritical, "FGBatchDetails"
    GoTo CleanExit

CleanExit:
    Set wsDest = Nothing
    Exit Sub
End Sub

Private Sub PopulateQuickCodeFromMaster(ByVal wsBatchDetails As Worksheet, ByVal lastRow As Long)
    On Error GoTo ErrHandler

    Dim wsMaster As Worksheet
    Dim i As Long
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim lookupKey As String
    Dim masterRow As Long
    Dim quickCode As String

    Set wsMaster = ThisWorkbook.Worksheets("FG_Master")

    For i = 2 To lastRow
        pmiNumber = wsBatchDetails.Cells(i, 2).value
        pmiVersion = wsBatchDetails.Cells(i, 3).value
        
        If pmiNumber <> "" And pmiVersion <> "" Then
            lookupKey = pmiNumber & "|" & pmiVersion
            masterRow = FindMasterRow(wsMaster, pmiNumber, pmiVersion)
            
            If masterRow > 0 Then
                quickCode = wsMaster.Cells(masterRow, 5).value
                wsBatchDetails.Cells(i, 17).value = quickCode
            End If
        End If
    Next i

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PopulateQuickCodeFromMaster"
    GoTo CleanExit

CleanExit:
    Set wsMaster = Nothing
    Exit Sub
End Sub

Private Function FindMasterRow(ByVal wsMaster As Worksheet, ByVal pmiCode As String, ByVal pmiVersion As String) As Long
    Dim i As Long
    Dim masterLastRow As Long

    masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).row

    For i = 2 To masterLastRow
        If wsMaster.Cells(i, 1).value = pmiCode And wsMaster.Cells(i, 2).value = pmiVersion Then
            FindMasterRow = i
            Exit Function
        End If
    Next i

    FindMasterRow = 0
End Function

Public Sub FGMasterDataLoad()
    On Error GoTo ErrHandler

    Dim fgMasterCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long
    Dim sourceFolderPath As String
    
    DisableAppUpdates

    sourceFolderPath = GetMasterDataTablesPath()
    fgMasterCSVPath = FILE_FG_MASTER_DATA
    Set wsDest = ThisWorkbook.Worksheets("FG_Master")

    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(fgMasterCSVPath, wsDest, wsDest.Range("A1"), True, sourceFolderPath)
    If Not importSuccess Then
        Err.Raise vbObjectError + 106, , ERROR_CSV_IMPORT & fgMasterCSVPath
    End If

    ' --- Update named ranges ---
    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    If lastRow >= 2 Then
        ThisWorkbook.Names("NRpmi_CodeMaster").RefersTo = wsDest.Range("A2:A" & lastRow)
        ThisWorkbook.Names("NRpmi_StatusMaster").RefersTo = wsDest.Range("G2:G" & lastRow)
        ThisWorkbook.Names("NRpmiProdDesc").RefersTo = wsDest.Range("F2:F" & lastRow)
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "FG Master Data Load Error: " & Err.Description, vbCritical, "FGMasterDataLoad"
    GoTo CleanExit

CleanExit:
    Set wsDest = Nothing
    Exit Sub
End Sub

Public Sub LabelPrintsDataLoad()
    On Error GoTo ErrHandler

    Dim labelPrintsCSVPath As String
    Dim wsDest As Worksheet
    Dim tblLabelPrintsImp As ListObject
    Dim importSuccess As Boolean
    Dim tempStartCell As Range
    Dim importedRange As Range
    Dim lastRow As Long
    Dim lastCol As Long

    DisableAppUpdates

    labelPrintsCSVPath = GetWorkbookTableFilePath(FILE_LABEL_DETAILS)
    Set wsDest = ThisWorkbook.Worksheets("Label_Prints")
    Set tblLabelPrintsImp = wsDest.ListObjects("tblLabelPrintsImp")

    If Not tblLabelPrintsImp.DataBodyRange Is Nothing Then
        tblLabelPrintsImp.DataBodyRange.Delete
    End If

    Set tempStartCell = wsDest.Range("Z1")
    
    importSuccess = ImportCSVWithOpenText(labelPrintsCSVPath, wsDest, tempStartCell, False)
    If Not importSuccess Then
        Err.Raise vbObjectError + 107, , ERROR_CSV_IMPORT & labelPrintsCSVPath
    End If

    lastRow = wsDest.Cells(wsDest.Rows.Count, tempStartCell.Column).End(xlUp).row
    lastCol = wsDest.Cells(tempStartCell.row, wsDest.Columns.Count).End(xlToLeft).Column
    
    Set importedRange = wsDest.Range(tempStartCell, wsDest.Cells(lastRow, lastCol))
    
    importedRange.Copy tblLabelPrintsImp.Range.Cells(1, 1)
    
    importedRange.EntireColumn.Delete
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Label Prints Data Load Error: " & Err.Description, vbCritical, "LabelPrintsDataLoad"
    GoTo CleanExit

CleanExit:
    Set importedRange = Nothing
    Set tempStartCell = Nothing
    Set tblLabelPrintsImp = Nothing
    Set wsDest = Nothing
    Exit Sub
End Sub

Public Sub LabelRegisterDataLoad()
    On Error GoTo ErrHandler

    Dim labelRegisterCSVPath As String
    Dim wsDest As Worksheet
    Dim tblLabelReg As ListObject
    Dim importSuccess As Boolean
    Dim tempStartCell As Range
    Dim importedRange As Range
    Dim lastRow As Long
    Dim lastCol As Long

    DisableAppUpdates

    labelRegisterCSVPath = GetWorkbookTableFilePath(FILE_LABEL_REGISTER)
    Set wsDest = ThisWorkbook.Worksheets("Label Register")
    Set tblLabelReg = wsDest.ListObjects("tblLabelReg")

    If Not tblLabelReg.DataBodyRange Is Nothing Then
        tblLabelReg.DataBodyRange.Delete
    End If

    Set tempStartCell = wsDest.Range("Z1")
    
    importSuccess = ImportCSVWithOpenText(labelRegisterCSVPath, wsDest, tempStartCell, False)
    If Not importSuccess Then
        Err.Raise vbObjectError + 108, , ERROR_CSV_IMPORT & labelRegisterCSVPath
    End If

    lastRow = wsDest.Cells(wsDest.Rows.Count, tempStartCell.Column).End(xlUp).row
    lastCol = wsDest.Cells(tempStartCell.row, wsDest.Columns.Count).End(xlToLeft).Column
    
    Set importedRange = wsDest.Range(tempStartCell, wsDest.Cells(lastRow, lastCol))
    
    importedRange.Copy tblLabelReg.Range.Cells(1, 1)
    
    importedRange.EntireColumn.Delete
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Label Register Data Load Error: " & Err.Description, vbCritical, "LabelRegisterDataLoad"
    GoTo CleanExit

CleanExit:
    Set importedRange = Nothing
    Set tempStartCell = Nothing
    Set tblLabelReg = Nothing
    Set wsDest = Nothing
    Exit Sub
End Sub

Public Sub LoadPrintReprintData( _
    ByVal qidCode As String, _
    ByVal cpmuRef As String, _
    ByVal dateFrom As Variant, _
    ByVal dateTo As Variant, _
    ByVal userName As String, _
    ByVal printedStatus As String, _
    ByVal batchFrom As String, _
    ByVal batchTo As String, _
    ByVal mfgLotFrom As String, _
    ByVal mfgLotTo As String, _
    ByVal mfgLot2From As String, _
    ByVal mfgLot2To As String)

    On Error GoTo ErrHandler

    Dim wsPrints As Worksheet
    Dim wsLabelPrints As Worksheet
    Dim tblPrints As ListObject
    Dim tblLabelPrintsImp As ListObject
    Dim fullTable As Range
    Dim dataArr As Variant
    Dim resultArr As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim writeRow As Long
    Dim outputCols() As Variant
    Dim colQID As Long
    Dim colBatch As Long
    Dim colMfgLot As Long
    Dim colMfgLot2 As Long
    Dim colCPMURef As Long
    Dim colCreatedBy As Long
    Dim colCreatedDate As Long
    Dim colLabelsPerSheet As Long
    Dim colLabelsRequested As Long
    Dim colTotalPrinted As Long
    Dim colProductDesc As Long
    Dim colPrintType As Long
    Dim colExpiry As Long
    Dim colPrintID As Long
    Dim colPrinted As Long
    Dim dateFromNum As Double
    Dim dateToNum As Double
    Dim curDate As Double
    Dim i As Long

    Set wsPrints = ThisWorkbook.Worksheets("Reprint Labels")
    Set wsLabelPrints = ThisWorkbook.Worksheets("Label_Prints")
    Set tblPrints = wsPrints.ListObjects("tblPrints")
    Set tblLabelPrintsImp = wsLabelPrints.ListObjects("tblLabelPrintsImp")

    FormUtils.UnprotectActiveSheet

    If Not tblPrints.DataBodyRange Is Nothing Then
        tblPrints.DataBodyRange.ClearContents
    End If

    If IsDate(dateFrom) Then dateFromNum = CDbl(CDate(dateFrom)) Else dateFromNum = 0
    If IsDate(dateTo) Then dateToNum = CDbl(CDate(dateTo)) + 0.9999 Else dateToNum = 0

    If (dateFromNum > 0 And dateToNum > 0) And (dateFromNum > dateToNum) Then
        GoTo CleanExit
    End If

    With tblLabelPrintsImp
        colQID = .ListColumns(LABEL_HEADER_QID_CODE).Index
        colBatch = .ListColumns(LABEL_HEADER_BATCH_NUMBER).Index
        colMfgLot = .ListColumns(LABEL_HEADER_MANUF_LOT_NO).Index
        colMfgLot2 = .ListColumns(LABEL_HEADER_MANUF_LOT_NO_2).Index
        colExpiry = .ListColumns(LABEL_HEADER_EXPIRY).Index
        colCPMURef = .ListColumns(LABEL_HEADER_PMI_NUMBER).Index
        colCreatedBy = .ListColumns(LABEL_HEADER_CREATED_BY).Index
        colCreatedDate = .ListColumns(LABEL_HEADER_CREATED_DATE).Index
        colLabelsPerSheet = .ListColumns(LABEL_HEADER_LABELS_PER_SHEET).Index
        colLabelsRequested = .ListColumns(LABEL_HEADER_QUANTITY).Index
        colTotalPrinted = .ListColumns(LABEL_HEADER_TOTAL_PRINTED).Index
        colProductDesc = .ListColumns(LABEL_HEADER_PRODUCT_DESCRIPTION).Index
        colPrinted = .ListColumns(LABEL_HEADER_PRINTED).Index
        colPrintID = .ListColumns(LABEL_HEADER_PRINT_ID).Index
        colPrintType = .ListColumns(LABEL_HEADER_PRINT_TYPE).Index
    End With

    Set fullTable = tblLabelPrintsImp.DataBodyRange
    dataArr = fullTable.value
    rowCount = UBound(dataArr, 1)
    colCount = UBound(dataArr, 2)

    outputCols = Array("Print ID", "Quick Code", "CPMU Ref", "Product Description", "Batch Number", "Mfg Lot No", "Mfg Lot No 2", "Expiry", "Labels per Sheet", "Labels Requested")
    ReDim resultArr(1 To rowCount, 1 To UBound(outputCols) + 1)
    writeRow = 1

        For i = 1 To rowCount
        If Len(qidCode) > 0 Then
            If dataArr(i, colQID) <> qidCode Then GoTo SkipRow
        End If
        
        If Len(cpmuRef) > 0 Then
            If dataArr(i, colCPMURef) <> cpmuRef Then GoTo SkipRow
        End If
        
        If Len(userName) > 0 Then
            If dataArr(i, colCreatedBy) <> userName Then GoTo SkipRow
        End If
        
        If Len(batchFrom) > 0 Then
            If dataArr(i, colBatch) < batchFrom Then GoTo SkipRow
        End If
        
        If Len(batchTo) > 0 Then
            If dataArr(i, colBatch) > batchTo Then GoTo SkipRow
        End If

        If IsDate(dataArr(i, colCreatedDate)) Then
            curDate = CDbl(CDate(dataArr(i, colCreatedDate)))
        Else
            curDate = 0
        End If
        
        If dateFromNum > 0 Then
            If curDate < dateFromNum Then GoTo SkipRow
        End If
        
        If dateToNum > 0 Then
            If curDate > dateToNum Then GoTo SkipRow
        End If

        If Len(mfgLotFrom) > 0 Then
            If dataArr(i, colMfgLot) < mfgLotFrom Then GoTo SkipRow
        End If
        
        If Len(mfgLotTo) > 0 Then
            If dataArr(i, colMfgLot) > mfgLotTo Then GoTo SkipRow
        End If
        
        If Len(mfgLot2From) > 0 Then
            If dataArr(i, colMfgLot2) < mfgLot2From Then GoTo SkipRow
        End If
        
        If Len(mfgLot2To) > 0 Then
            If dataArr(i, colMfgLot2) > mfgLot2To Then GoTo SkipRow
        End If

        If Len(dataArr(i, colPrintType)) > 0 Then
            If dataArr(i, colPrintType) <> "Original" Then GoTo SkipRow
        End If

        If Len(printedStatus) > 0 Then
            If printedStatus = "Yes" And dataArr(i, colPrinted) <> "x" Then GoTo SkipRow
            If printedStatus = "No" And dataArr(i, colPrinted) = "x" Then GoTo SkipRow
        End If

        resultArr(writeRow, 1) = dataArr(i, colPrintID)
        resultArr(writeRow, 2) = dataArr(i, colQID)
        resultArr(writeRow, 3) = dataArr(i, colCPMURef)
        resultArr(writeRow, 4) = dataArr(i, colProductDesc)
        resultArr(writeRow, 5) = dataArr(i, colBatch)
        resultArr(writeRow, 6) = dataArr(i, colMfgLot)
        resultArr(writeRow, 7) = dataArr(i, colMfgLot2)
        resultArr(writeRow, 8) = dataArr(i, colExpiry)
        resultArr(writeRow, 9) = dataArr(i, colLabelsPerSheet)
        resultArr(writeRow, 10) = dataArr(i, colLabelsRequested)

        writeRow = writeRow + 1

SkipRow:
    Next i

    If writeRow > 1 Then
        QuickSort2D resultArr, 1, writeRow - 1, 1
        wsPrints.Range("B12").Resize(writeRow - 1, UBound(outputCols) + 1).value = resultArr
        tblPrints.Resize wsPrints.Range("B11").Resize(writeRow, UBound(outputCols) + 5)
    Else
        MsgBox "No matching records found.", vbExclamation
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error in LoadPrintReprintData: " & Err.Description, vbCritical, "LoadPrintReprintData"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    FormUtils.ProtectActiveSheet
    wsPrints.Cells.Locked = True

    If Not tblPrints.DataBodyRange Is Nothing Then
        tblPrints.DataBodyRange.Locked = False
    End If

    If Not tblPrints.HeaderRowRange Is Nothing Then
        tblPrints.HeaderRowRange.Locked = False
    End If

    wsPrints.Range("C5:C6, F5:F6, H5, K5:K7, M5:M7").Locked = False
    tblPrints.ShowAutoFilter = True

    If PROD_DEV <> "DEV" Then
        wsPrints.Protect , _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            UserInterfaceOnly:=True
    End If

    wsPrints.EnableAutoFilter = True

    Set wsPrints = Nothing
    Set wsLabelPrints = Nothing
    Set tblPrints = Nothing
    Set tblLabelPrintsImp = Nothing
    Set fullTable = Nothing
    Exit Sub
End Sub

Public Sub LoadPrintData( _
    ByVal qidCode As String, _
    ByVal cpmuRef As String, _
    ByVal dateFrom As Variant, _
    ByVal dateTo As Variant, _
    ByVal userName As String, _
    ByVal batchFrom As String, _
    ByVal batchTo As String, _
    ByVal mfgLotFrom As String, _
    ByVal mfgLotTo As String, _
    ByVal mfgLot2From As String, _
    ByVal mfgLot2To As String)

    On Error GoTo ErrHandler

    Dim wsPrevious As Worksheet
    Dim wsLabelPrints As Worksheet
    Dim tblPrevious As ListObject
    Dim tblLabelPrintsImp As ListObject
    Dim fullTable As Range
    Dim dataArr As Variant
    Dim resultArr As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim writeRow As Long
    Dim outputCols() As Variant
    Dim colQID As Long
    Dim colBatch As Long
    Dim colMfgLot As Long
    Dim colMfgLot2 As Long
    Dim colCPMURef As Long
    Dim colPrintedBy As Long
    Dim colPrintDate As Long
    Dim colLabelsPerSheet As Long
    Dim colLabelsRequested As Long
    Dim colTotalPrinted As Long
    Dim colProductDesc As Long
    Dim colPrintType As Long
    Dim colReprintReason As Long
    Dim colExpiry As Long
    Dim colPrintID As Long
    Dim dateFromNum As Double
    Dim dateToNum As Double
    Dim curDate As Double
    Dim i As Long

    Set wsPrevious = ThisWorkbook.Worksheets("Previous Prints")
    Set wsLabelPrints = ThisWorkbook.Worksheets("Label_Prints")
    Set tblPrevious = wsPrevious.ListObjects("tblPreviousPrints")
    Set tblLabelPrintsImp = wsLabelPrints.ListObjects("tblLabelPrintsImp")

    FormUtils.UnprotectActiveSheet

    If Not tblPrevious.DataBodyRange Is Nothing Then
        tblPrevious.DataBodyRange.ClearContents
    End If

    If IsDate(dateFrom) Then dateFromNum = CDbl(CDate(dateFrom)) Else dateFromNum = 0
    If IsDate(dateTo) Then dateToNum = CDbl(CDate(dateTo)) + 0.9999 Else dateToNum = 0

    If (dateFromNum > 0 And dateToNum > 0) And (dateFromNum > dateToNum) Then
        GoTo CleanExit
    End If

    With tblLabelPrintsImp
        colPrintID = .ListColumns(LABEL_HEADER_PRINT_ID).Index
        colQID = .ListColumns(LABEL_HEADER_QID_CODE).Index
        colCPMURef = .ListColumns(LABEL_HEADER_PMI_NUMBER).Index
        colProductDesc = .ListColumns(LABEL_HEADER_PRODUCT_DESCRIPTION).Index
        colBatch = .ListColumns(LABEL_HEADER_BATCH_NUMBER).Index
        colExpiry = .ListColumns(LABEL_HEADER_EXPIRY).Index
        colMfgLot = .ListColumns(LABEL_HEADER_MANUF_LOT_NO).Index
        colMfgLot2 = .ListColumns(LABEL_HEADER_MANUF_LOT_NO_2).Index
        colLabelsPerSheet = .ListColumns(LABEL_HEADER_LABELS_PER_SHEET).Index
        colLabelsRequested = .ListColumns(LABEL_HEADER_QUANTITY).Index
        colTotalPrinted = .ListColumns(LABEL_HEADER_TOTAL_PRINTED).Index
        colPrintType = .ListColumns(LABEL_HEADER_PRINT_TYPE).Index
        colReprintReason = .ListColumns(LABEL_HEADER_REPRINT_REASON).Index
        colPrintedBy = .ListColumns(LABEL_HEADER_CREATED_BY).Index
        colPrintDate = .ListColumns(LABEL_HEADER_CREATED_DATE).Index
    End With

    Set fullTable = tblLabelPrintsImp.DataBodyRange
    dataArr = fullTable.value
    rowCount = UBound(dataArr, 1)
    colCount = UBound(dataArr, 2)

    outputCols = Array("Print ID", "Quick Code", "CPMU Ref", "Product Description", "Batch Number", "Expiry", "Mfg Lot No", "Mfg Lot No 2", "Labels per Sheet", "Labels Requested", "Total Printed", "Print Type", "Reprint Reason", "Printed By", "Printed Date")
    ReDim resultArr(1 To rowCount, 1 To UBound(outputCols) + 1)
    writeRow = 1

    For i = 1 To rowCount
        If Len(qidCode) > 0 Then
            If dataArr(i, colQID) <> qidCode Then GoTo SkipRow
        End If
        
        If Len(cpmuRef) > 0 Then
            If dataArr(i, colCPMURef) <> cpmuRef Then GoTo SkipRow
        End If
        
        If Len(userName) > 0 Then
            If dataArr(i, colPrintedBy) <> userName Then GoTo SkipRow
        End If
        
        If Len(batchFrom) > 0 Then
            If dataArr(i, colBatch) < batchFrom Then GoTo SkipRow
        End If
        
        If Len(batchTo) > 0 Then
            If dataArr(i, colBatch) > batchTo Then GoTo SkipRow
        End If

        If IsDate(dataArr(i, colPrintDate)) Then
            curDate = CDbl(CDate(dataArr(i, colPrintDate)))
        Else
            curDate = 0
        End If
        
        If dateFromNum > 0 Then
            If curDate < dateFromNum Then GoTo SkipRow
        End If
        
        If dateToNum > 0 Then
            If curDate > dateToNum Then GoTo SkipRow
        End If
        
        If Len(mfgLotFrom) > 0 Then
            If dataArr(i, colMfgLot) < mfgLotFrom Then GoTo SkipRow
        End If
        
        If Len(mfgLotTo) > 0 Then
            If dataArr(i, colMfgLot) > mfgLotTo Then GoTo SkipRow
        End If
        
        If Len(mfgLot2From) > 0 Then
            If dataArr(i, colMfgLot2) < mfgLot2From Then GoTo SkipRow
        End If
        
        If Len(mfgLot2To) > 0 Then
            If dataArr(i, colMfgLot2) > mfgLot2To Then GoTo SkipRow
        End If

        resultArr(writeRow, 1) = dataArr(i, colPrintID)
        resultArr(writeRow, 2) = dataArr(i, colQID)
        resultArr(writeRow, 3) = dataArr(i, colCPMURef)
        resultArr(writeRow, 4) = dataArr(i, colProductDesc)
        resultArr(writeRow, 5) = dataArr(i, colBatch)
        resultArr(writeRow, 6) = dataArr(i, colExpiry)
        resultArr(writeRow, 7) = dataArr(i, colMfgLot)
        resultArr(writeRow, 8) = dataArr(i, colMfgLot2)
        resultArr(writeRow, 9) = dataArr(i, colLabelsPerSheet)
        resultArr(writeRow, 10) = dataArr(i, colLabelsRequested)
        resultArr(writeRow, 11) = dataArr(i, colTotalPrinted)
        resultArr(writeRow, 12) = dataArr(i, colPrintType)
        resultArr(writeRow, 13) = dataArr(i, colReprintReason)
        resultArr(writeRow, 14) = dataArr(i, colPrintedBy)
        resultArr(writeRow, 15) = dataArr(i, colPrintDate)

        writeRow = writeRow + 1

SkipRow:
    Next i

    If writeRow > 1 Then
        wsPrevious.Range("B12").Resize(writeRow - 1, UBound(outputCols) + 1).value = resultArr
        tblPrevious.Resize wsPrevious.Range("B11").Resize(writeRow, UBound(outputCols) + 2)
        wsPrevious.Range("G12").Resize(writeRow - 1, 1).NumberFormat = "dd/mm/yyyy"
    Else
        MsgBox "No matching records found.", vbExclamation
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error in LoadPrintData: " & Err.Description, vbCritical, "LoadPrintData"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    FormUtils.UnprotectActiveSheet
    wsPrevious.Cells.Locked = True

    If Not tblPrevious.DataBodyRange Is Nothing Then
        tblPrevious.DataBodyRange.Locked = False
    End If
    
    If Not tblPrevious.HeaderRowRange Is Nothing Then
        tblPrevious.HeaderRowRange.Locked = False
    End If
    
    wsPrevious.Range("D5:D6, G5:G8, I6:I8").Locked = False
    tblPrevious.ShowAutoFilter = True

    If PROD_DEV <> "DEV" Then
        wsPrevious.Protect , _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            UserInterfaceOnly:=True
    End If
        
    wsPrevious.EnableAutoFilter = True

    Set fullTable = Nothing
    Set tblLabelPrintsImp = Nothing
    Set tblPrevious = Nothing
    Set wsLabelPrints = Nothing
    Set wsPrevious = Nothing
    Exit Sub
End Sub

Public Sub LoadPMIViewerData( _
    ByVal qidCode As String, _
    ByVal pmiNumber As String, _
    ByVal batchNumber As String, _
    ByVal productDescription As String)

    On Error GoTo ErrHandler

    Dim wsBatchViewer As Worksheet
    Dim wsBatchDetails As Worksheet
    Dim wsFGMaster As Worksheet
    Dim tblPMIViewer As ListObject
    Dim fullTable As Range
    Dim fgMasterData As Variant
    Dim dataArr As Variant
    Dim resultArr As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim writeRow As Long
    Dim colPMI As Long
    Dim colProductDesc As Long
    Dim colBatch As Long
    Dim colExpiry As Long
    Dim colBatchStatus As Long
    Dim colScanned As Long
    Dim colFGMasterPMI As Long
    Dim colFGMasterQID As Long
    Dim colFGMasterStatus As Long
    Dim fgMasterRowCount As Long
    Dim lookupQID As String
    Dim currentPMI As String
    Dim batchStatus As String
    Dim i As Long
    Dim j As Long

    Set wsBatchViewer = ThisWorkbook.Worksheets("Batch Doc Viewer")
    Set wsBatchDetails = ThisWorkbook.Worksheets("Batch_Details")
    Set wsFGMaster = ThisWorkbook.Worksheets("FG_Master")
    Set tblPMIViewer = wsBatchViewer.ListObjects("tblPMIViewer")

    FormUtils.UnprotectActiveSheet

    If Not tblPMIViewer.DataBodyRange Is Nothing Then
        tblPMIViewer.DataBodyRange.ClearContents
    End If

    If wsBatchDetails.Cells(2, 1).value = "" Then
        GoTo CleanExit
    End If

    Set fullTable = wsBatchDetails.Range("A2").CurrentRegion
    If fullTable.Rows.Count < 2 Then
        GoTo CleanExit
    End If
    
    dataArr = fullTable.value
    rowCount = UBound(dataArr, 1)
    colCount = UBound(dataArr, 2)

    fgMasterData = wsFGMaster.Range("A2").CurrentRegion.value
    fgMasterRowCount = UBound(fgMasterData, 1)

    colBatch = 1
    colPMI = 2
    colProductDesc = 4
    colExpiry = 8
    colBatchStatus = 11
    colScanned = 16

    colFGMasterPMI = 1
    colFGMasterQID = 5
    colFGMasterStatus = 7

    ReDim resultArr(1 To rowCount - 1, 1 To 6)
    writeRow = 1

    For i = 2 To rowCount
        batchStatus = Trim(CStr(dataArr(i, colBatchStatus)))
        
        If batchStatus = "Created" Or batchStatus = "Cancelled" Then
            GoTo SkipRow
        End If
        
        If Len(batchNumber) > 0 Then
            If dataArr(i, colBatch) <> batchNumber Then GoTo SkipRow
        End If
        
        If Len(productDescription) > 0 Then
            If dataArr(i, colProductDesc) <> productDescription Then GoTo SkipRow
        End If
        
        currentPMI = dataArr(i, colPMI)
        
        If Len(pmiNumber) > 0 Then
            If currentPMI <> pmiNumber Then GoTo SkipRow
        End If

        lookupQID = ""
        For j = 1 To fgMasterRowCount
            If fgMasterData(j, colFGMasterPMI) = currentPMI And _
               fgMasterData(j, colFGMasterStatus) = "Current" Then
                lookupQID = fgMasterData(j, colFGMasterQID)
                Exit For
            End If
        Next j

        If Len(qidCode) > 0 Then
            If lookupQID <> qidCode Then GoTo SkipRow
        End If

        resultArr(writeRow, 1) = lookupQID
        resultArr(writeRow, 2) = currentPMI
        resultArr(writeRow, 3) = dataArr(i, colProductDesc)
        resultArr(writeRow, 4) = dataArr(i, colBatch)
        resultArr(writeRow, 5) = dataArr(i, colExpiry)
        resultArr(writeRow, 6) = dataArr(i, colScanned)

        writeRow = writeRow + 1

SkipRow:
    Next i

    If writeRow > 1 Then
        wsBatchViewer.Range("B12").Resize(writeRow - 1, 6).value = resultArr
        tblPMIViewer.Resize wsBatchViewer.Range("B11").Resize(writeRow, 7)
        wsBatchViewer.Range("F12").Resize(writeRow - 1, 1).NumberFormat = "dd/mm/yyyy"
    Else
        MsgBox "No matching records found.", vbExclamation
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error in LoadPMIViewerData: " & Err.Description, vbCritical, "LoadPMIViewerData"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    FormUtils.ProtectActiveSheet
    wsBatchViewer.Cells.Locked = True

    If Not tblPMIViewer.DataBodyRange Is Nothing Then
        tblPMIViewer.DataBodyRange.Locked = False
    End If
    
    If Not tblPMIViewer.HeaderRowRange Is Nothing Then
        tblPMIViewer.HeaderRowRange.Locked = False
    End If
    
    wsBatchViewer.Range("D5:D8").Locked = False
    tblPMIViewer.ShowAutoFilter = True

    If PROD_DEV <> "DEV" Then
        wsBatchViewer.Protect , _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            UserInterfaceOnly:=True
    End If
        
    wsBatchViewer.EnableAutoFilter = True

    Set fullTable = Nothing
    Set tblPMIViewer = Nothing
    Set wsFGMaster = Nothing
    Set wsBatchDetails = Nothing
    Set wsBatchViewer = Nothing
    Exit Sub
End Sub