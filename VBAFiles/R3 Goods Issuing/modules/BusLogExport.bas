Attribute VB_Name = "BusLogExport"
' ========================================================================
' WHERE USED REPORT GENERATION
' ========================================================================
' Generates a "Where Used" report for QC Numbers
' Traces QC Numbers through RM_Trans to Finished Goods batches
' ========================================================================

Option Explicit

Public Sub GenerateWhereUsedReport()
    On Error GoTo ErrorHandler
    
    Dim qcCSVPath As String
    Dim rmTransPath As String
    Dim batchMasterPath As String
    Dim qcData As Variant
    Dim rmTransData As Variant
    Dim batchData As Variant
    Dim filteredQCNumbers As Collection
    Dim outputData As Collection
    Dim i As Long
    Dim filterQCNumber As String
    Dim filterStdName As String
    Dim filterRMCode As String
    Dim filterVendor As String
    Dim filterMfgName As String
    Dim filterMfgLot As String
    
    filterQCNumber = Trim(GetFieldValByLabelC("QC Number:"))
    filterStdName = Trim(GetFieldValByLabelC("Standard Name:"))
    filterRMCode = RM_Pad4(Trim(GetFieldValByLabelC("RM Code:")))
    filterVendor = Trim(GetFieldValByLabelC("Vendor Name:"))
    filterMfgName = Trim(GetFieldValByLabelC("Manufacturer Name:"))
    filterMfgLot = Trim(GetFieldValByLabelC("Manufacturer Lot Number:"))
    
    ' Validate that at least one filter is provided
    If filterQCNumber = "" And filterStdName = "" And filterRMCode = "" And _
       filterVendor = "" And filterMfgName = "" And filterMfgLot = "" Then
        MsgBox "Please provide at least one filter criterion.", vbExclamation, "No Filter"
        GoTo CleanExit
    End If
    
    ' Build file paths
    qcCSVPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    rmTransPath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    batchMasterPath = GetWorkbookTableFilePath(FILE_FINISHED_GOODS_BATCH_MASTER)
    
    ' Validate files exist
    If Not FileExists(qcCSVPath) Then
        MsgBox "QC_Number.csv file not found at: " & qcCSVPath, vbCritical, "File Missing"
        GoTo CleanExit
    End If
    If Not FileExists(rmTransPath) Then
        MsgBox "RM_Trans.csv file not found at: " & rmTransPath, vbCritical, "File Missing"
        GoTo CleanExit
    End If
    If Not FileExists(batchMasterPath) Then
        MsgBox "FinishedGoodsBatchNumberMasterFile.csv file not found at: " & batchMasterPath, vbCritical, "File Missing"
        GoTo CleanExit
    End If
    
    ' Load CSV data
    qcData = LoadCSVToArray(qcCSVPath)
    rmTransData = LoadCSVToArray(rmTransPath)
    batchData = LoadCSVToArray(batchMasterPath)
    
    If IsEmpty(qcData) Or IsEmpty(rmTransData) Or IsEmpty(batchData) Then
        MsgBox "Failed to load one or more CSV files.", vbCritical, "Load Error"
        GoTo CleanExit
    End If
    
    ' Step 1: Filter QC Numbers based on criteria
    Set filteredQCNumbers = FilterQCNumbers(qcData, filterQCNumber, filterStdName, _
                                            filterRMCode, filterVendor, filterMfgName, filterMfgLot)
    
    If filteredQCNumbers.Count = 0 Then
        MsgBox "No QC Numbers found matching the specified criteria.", vbInformation, "No Results"
        GoTo CleanExit
    End If
    
    ' Step 2: Find FG Batches where filtered QC Numbers were used
    ' Step 3: Get batch details excluding reprints
    Set outputData = BuildWhereUsedData(filteredQCNumbers, qcData, rmTransData, batchData)
    
    If outputData.Count = 0 Then
        MsgBox "No finished goods batches found for the selected QC Numbers.", vbInformation, "No Results"
        GoTo CleanExit
    End If
    
    ' Step 4: Output results to a new workbook
    OutputWhereUsedReport outputData
    
    MsgBox "Where Used report generated successfully." & vbCrLf & _
           "Total records: " & outputData.Count, vbInformation, "Success"
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error generating Where Used report: " & Err.Description, vbCritical, "GenerateWhereUsedReport"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Set filteredQCNumbers = Nothing
    Set outputData = Nothing
    On Error GoTo 0
    Exit Sub
End Sub

' Filters QC Numbers based on provided criteria
Private Function FilterQCNumbers(ByVal qcData As Variant, _
                                 ByVal filterQCNumber As String, _
                                 ByVal filterStdName As String, _
                                 ByVal filterRMCode As String, _
                                 ByVal filterVendor As String, _
                                 ByVal filterMfgName As String, _
                                 ByVal filterMfgLot As String) As Collection
    Dim results As Collection
    Dim i As Long
    Dim colIdx As Object
    Dim qcNum As String
    Dim stdName As String
    Dim rmCode As String
    Dim vendor As String
    Dim mfgName As String
    Dim mfgLot As String
    Dim matchFound As Boolean
    
    Set results = New Collection
    Set colIdx = GetColumnIndexes(qcData)
    
    ' Loop through data rows (skip header)
    For i = LBound(qcData, 1) + 1 To UBound(qcData, 1)
        matchFound = True
        
        ' Check each filter criterion
        If filterQCNumber <> "" Then
            qcNum = CStr(qcData(i, colIdx(HEADER_QC_NUMBER)))
            If InStr(1, qcNum, filterQCNumber, vbTextCompare) = 0 Then
                matchFound = False
            End If
        End If
        
        If matchFound And filterStdName <> "" Then
            stdName = CStr(qcData(i, colIdx(HEADER_STD_NAME)))
            If InStr(1, stdName, filterStdName, vbTextCompare) = 0 Then
                matchFound = False
            End If
        End If
        
        If matchFound And filterRMCode <> "" Then
            rmCode = RM_Pad4(CStr(qcData(i, colIdx(HEADER_RM_CODE))))
            If InStr(1, rmCode, filterRMCode, vbTextCompare) = 0 Then
                matchFound = False
            End If
        End If
        
        If matchFound And filterVendor <> "" Then
            vendor = CStr(qcData(i, colIdx(HEADER_VENDOR_NAME)))
            If InStr(1, vendor, filterVendor, vbTextCompare) = 0 Then
                matchFound = False
            End If
        End If
        
        If matchFound And filterMfgName <> "" Then
            mfgName = CStr(qcData(i, colIdx(HEADER_MANUFACTURER_NAME)))
            If InStr(1, mfgName, filterMfgName, vbTextCompare) = 0 Then
                matchFound = False
            End If
        End If
        
        If matchFound And filterMfgLot <> "" Then
            mfgLot = CStr(qcData(i, colIdx(HEADER_MANUFACTURER_LOT)))
            If InStr(1, mfgLot, filterMfgLot, vbTextCompare) = 0 Then
                matchFound = False
            End If
        End If
        
        ' If all filters match, add QC Number to results
        If matchFound Then
            qcNum = CStr(qcData(i, colIdx(HEADER_QC_NUMBER)))
            On Error Resume Next
            results.Add qcNum, qcNum
            On Error GoTo 0
        End If
    Next i
    
    Set FilterQCNumbers = results
End Function

' Builds the where-used data by joining QC Numbers -> RM_Trans -> Batch Details
Private Function BuildWhereUsedData(ByVal qcNumbers As Collection, _
                                    ByVal qcData As Variant, _
                                    ByVal rmTransData As Variant, _
                                    ByVal batchData As Variant) As Collection
    On Error GoTo ErrorHandler
    
    Dim results As Collection
    Dim qcColIdx As Object
    Dim rmColIdx As Object
    Dim batchColIdx As Object
    Dim batchQtyDict As Object
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim qcNum As String
    Dim fgBatch As String
    Dim batchNum As String
    Dim transType As String
    Dim reversedFlag As String
    Dim fgBatches As Collection
    Dim outputRow As Variant
    Dim qtyUsed As Double
    Dim masterUoM As String
    Dim mfgLot As String
    Dim mfgExpiry As String
    Dim qcRow As Long
    Dim batchRow As Long
    
    Set results = New Collection
    Set qcColIdx = GetColumnIndexes(qcData)
    Set rmColIdx = GetColumnIndexes(rmTransData)
    Set batchColIdx = GetColumnIndexes(batchData)
    
    ' For each filtered QC Number
    For i = 1 To qcNumbers.Count
        qcNum = qcNumbers(i)
        
        ' Find row in QC data for this QC Number
        qcRow = FindQCRow(qcData, qcColIdx, qcNum)
        
        If qcRow = 0 Then
            GoTo NextQCNumber
        End If
        
        ' Get Mfg_Lot and Mfg_Expiry from QC data
        mfgLot = CStr(qcData(qcRow, qcColIdx(HEADER_MANUFACTURER_LOT)))
        mfgExpiry = CStr(qcData(qcRow, qcColIdx(HEADER_MANUFACTURER_EXPIRY)))
        masterUoM = CStr(qcData(qcRow, qcColIdx(HEADER_MASTER_UOM)))
        
        ' Build dictionary to sum quantities by FG Batch for CON transactions (excluding reversed)
        Set batchQtyDict = CreateObject("Scripting.Dictionary")
        Set fgBatches = New Collection
        
        For j = LBound(rmTransData, 1) + 1 To UBound(rmTransData, 1)
            If CStr(rmTransData(j, rmColIdx(RMT_HEADER_QC_NUMBER))) = qcNum Then
                transType = Trim(CStr(rmTransData(j, rmColIdx(RMT_HEADER_TRANS_TYPE))))
                reversedFlag = Trim(CStr(rmTransData(j, rmColIdx(RMT_HEADER_REVERSED_FLAG))))
                
                If UCase$(transType) = UCase$(TRANS_TYPE_CONSUMED) And UCase$(reversedFlag) <> "Y" Then
                    fgBatch = Trim(CStr(rmTransData(j, rmColIdx(RMT_HEADER_FG_BATCH))))
                    
                    If fgBatch <> "" Then
                        If batchQtyDict.Exists(fgBatch) Then
                            batchQtyDict(fgBatch) = batchQtyDict(fgBatch) + CDbl(rmTransData(j, rmColIdx(RMT_HEADER_QUANTITY)))
                        Else
                            batchQtyDict.Add fgBatch, CDbl(rmTransData(j, rmColIdx(RMT_HEADER_QUANTITY)))
                            On Error Resume Next
                            fgBatches.Add fgBatch, fgBatch
                            On Error GoTo ErrorHandler
                        End If
                    End If
                End If
            End If
        Next j
        
        ' For each FG Batch, find details in Label_Prints (use LAST entry for batch)
        For j = 1 To fgBatches.Count
            fgBatch = fgBatches(j)
            qtyUsed = Abs(batchQtyDict(fgBatch))
            
            If qtyUsed > 0 Then
                batchRow = FindBatchRow(batchData, batchColIdx, fgBatch)
                
                If batchRow > 0 Then
                    ReDim outputRow(1 To 14)
                    outputRow(1) = qcNum
                    outputRow(2) = CStr(qcData(qcRow, qcColIdx(HEADER_STD_NAME)))
                    outputRow(3) = CStr(qcData(qcRow, qcColIdx(HEADER_RM_CODE)))
                    outputRow(4) = mfgLot
                    outputRow(5) = mfgExpiry
                    outputRow(6) = qtyUsed
                    outputRow(7) = masterUoM
                    outputRow(8) = batchData(batchRow, batchColIdx(FGBNM_HEADER_BATCH_NUMBER))
                    outputRow(9) = batchData(batchRow, batchColIdx(FGBNM_HEADER_PMI_NUMBER))
                    outputRow(10) = batchData(batchRow, batchColIdx(FGBNM_HEADER_PMI_VERSION))
                    outputRow(11) = batchData(batchRow, batchColIdx(FGBNM_HEADER_PRODUCT_DESCRIPTION))
                    outputRow(12) = batchData(batchRow, batchColIdx(FGBNM_HEADER_MANUFACTURING_DATE))
                    outputRow(13) = batchData(batchRow, batchColIdx(FGBNM_HEADER_EXPIRY_DATE))
                    outputRow(14) = batchData(batchRow, batchColIdx(FGBNM_HEADER_BATCH_STATUS))
                    
                    results.Add outputRow
                End If
            End If
        Next j
        
NextQCNumber:
    Next i
    
    Set BuildWhereUsedData = results
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in BuildWhereUsedData: " & Err.Number & " - " & Err.Description, vbCritical, "BuildWhereUsedData"
    Set BuildWhereUsedData = New Collection
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Set fgBatches = Nothing
    Set batchQtyDict = Nothing
    Set qcColIdx = Nothing
    Set rmColIdx = Nothing
    Set batchColIdx = Nothing
    On Error GoTo 0
    Exit Function
End Function

' Helper: Find row index for a given QC Number
Private Function FindQCRow(ByVal qcData As Variant, ByVal colIdx As Object, ByVal qcNum As String) As Long
    Dim i As Long
    
    For i = LBound(qcData, 1) + 1 To UBound(qcData, 1)
        If CStr(qcData(i, colIdx(HEADER_QC_NUMBER))) = qcNum Then
            FindQCRow = i
            Exit Function
        End If
    Next i
    
    FindQCRow = 0
End Function

Private Function FindBatchRow(ByVal batchData As Variant, ByVal colIdx As Object, ByVal batchNum As String) As Long
    Dim i As Long
    
    For i = LBound(batchData, 1) + 1 To UBound(batchData, 1)
        If CStr(batchData(i, colIdx(FGBNM_HEADER_BATCH_NUMBER))) = batchNum Then
            FindBatchRow = i
        End If
    Next i
End Function

Private Sub OutputWhereUsedReport(ByVal data As Collection)
    On Error GoTo ErrorHandler
    
    Dim newWb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim outputRow As Variant
    Dim newRow(1 To 14) As Variant
    
    Set newWb = Workbooks.Add
    Set ws = newWb.Worksheets(1)
    ws.Name = "WhereUsed_Report"
    
    ws.Cells(1, 1).value = "RM_Code"
    ws.Cells(1, 2).value = "Standard_Name"
    ws.Cells(1, 3).value = "QC_Number"
    ws.Cells(1, 4).value = "Mfg_Lot"
    ws.Cells(1, 5).value = "Mfg_Expiry"
    ws.Cells(1, 6).value = "Quantity_Used"
    ws.Cells(1, 7).value = "Master_UoM"
    ws.Cells(1, 8).value = "Batch_Number"
    ws.Cells(1, 9).value = "PMI_Number"
    ws.Cells(1, 10).value = "PMI_Version"
    ws.Cells(1, 11).value = "Product_Description"
    ws.Cells(1, 12).value = "Manufacturing_Date"
    ws.Cells(1, 13).value = "Expiry_Date"
    ws.Cells(1, 14).value = "Batch_Status"

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 7))
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230)
    End With

    With ws.Range(ws.Cells(1, 8), ws.Cells(1, 14))
        .Font.Bold = True
        .Interior.Color = RGB(144, 238, 144)
    End With
    
    ' Write data with column mapping
    For i = 1 To data.Count
        outputRow = data(i)
        
        newRow(1) = outputRow(3)
        newRow(2) = outputRow(2)
        newRow(3) = outputRow(1)
        newRow(4) = outputRow(4)
        newRow(5) = outputRow(5)
        newRow(6) = outputRow(6)
        newRow(7) = outputRow(7)
        newRow(8) = outputRow(8)
        newRow(9) = outputRow(9)
        newRow(10) = outputRow(10)
        newRow(11) = outputRow(11)
        newRow(12) = outputRow(12)
        newRow(13) = outputRow(13)
        newRow(14) = outputRow(14)
        
        For j = 1 To 14
            If j = 1 Then
                With ws.Cells(i + 1, j)
                    .NumberFormat = "@"
                    .value = RM_Pad4(CStr(newRow(j)))
                End With
            Else
                ws.Cells(i + 1, j).value = newRow(j)
            End If
        Next j
    Next i
    
    ws.Columns.AutoFit
    newWb.Windows(1).Visible = True
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in OutputWhereUsedReport: " & Err.Number & " - " & Err.Description, vbCritical, "OutputWhereUsedReport"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Set ws = Nothing
    Set newWb = Nothing
    On Error GoTo 0
    Exit Sub
End Sub

' Helper: Creates a dictionary of column names to indexes
Private Function GetColumnIndexes(ByVal dataArray As Variant) As Object
    Dim dict As Object
    Dim i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    
    For i = LBound(dataArray, 2) To UBound(dataArray, 2)
        dict(Trim(CStr(dataArray(LBound(dataArray, 1), i)))) = i
    Next i
    
    Set GetColumnIndexes = dict
End Function

' Helper: Loads CSV file into a 2D array
Private Function LoadCSVToArray(ByVal filePath As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim tempWb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataArray As Variant
    
    Set tempWb = Workbooks.Open(filePath, , True)
    Set ws = tempWb.Worksheets(1)
    
    ' Get data range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 0 And lastCol > 0 Then
        dataArray = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).value
    End If
    
    ' Close the temporary workbook
    tempWb.Close SaveChanges:=False
    
    LoadCSVToArray = dataArray
    GoTo CleanExit
    
ErrorHandler:
    LoadCSVToArray = Empty
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Set ws = Nothing
    If Not tempWb Is Nothing Then
        tempWb.Close SaveChanges:=False
        Set tempWb = Nothing
    End If
    On Error GoTo 0
    Exit Function
End Function
