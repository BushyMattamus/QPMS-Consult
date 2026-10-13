Attribute VB_Name = "BusLogReportingExpiryProfile"
' ========================================================================
' RETEST ANALYSIS REPORTING
' ========================================================================
' Generates retest due date analysis from QC_Number worksheet data
' Categorizes QC numbers by retest urgency (Past Due, 1/3/6 months)
' Displays eligible items with balance and retest date information
' Also generates manufacturer expiry analysis for tested raw materials
' ========================================================================

Option Explicit

Public Sub GenerateRetestAnalysisReport()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim wsQCData As Worksheet
    Dim pastDueItems As Object
    Dim oneMonthItems As Object
    Dim threeMonthItems As Object
    Dim sixMonthItems As Object
    Dim outputRow As Long
    
    LoadQCNumberData
    
    Set ws = ActiveSheet
    Set wsQCData = ThisWorkbook.Worksheets("QC_Number")
    Set pastDueItems = CreateObject("Scripting.Dictionary")
    Set oneMonthItems = CreateObject("Scripting.Dictionary")
    Set threeMonthItems = CreateObject("Scripting.Dictionary")
    Set sixMonthItems = CreateObject("Scripting.Dictionary")
    
    CategorizeRetestItems wsQCData, pastDueItems, oneMonthItems, threeMonthItems, sixMonthItems
    
    UnprotectActiveSheet
    
    ws.Range("C6").value = pastDueItems.count
    ws.Range("D6").value = oneMonthItems.count
    ws.Range("E6").value = threeMonthItems.count
    ws.Range("F6").value = sixMonthItems.count
    
    ws.Range("H10:N10000").ClearContents
    
    outputRow = 10
    OutputRetestReportData ws, "Past Due", pastDueItems, outputRow
    OutputRetestReportData ws, "1 Month", oneMonthItems, outputRow
    OutputRetestReportData ws, "3 Month", threeMonthItems, outputRow
    OutputRetestReportData ws, "6 Month", sixMonthItems, outputRow
    
    ProtectActiveSheet
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateRetestAnalysisReport"
    GoTo CleanExit

CleanExit:
    Set ws = Nothing
    Set wsQCData = Nothing
    Set pastDueItems = Nothing
    Set oneMonthItems = Nothing
    Set threeMonthItems = Nothing
    Set sixMonthItems = Nothing
    Exit Sub
End Sub

Private Sub CategorizeRetestItems( _
    ByVal wsQCData As Worksheet, _
    ByVal pastDueItems As Object, _
    ByVal oneMonthItems As Object, _
    ByVal threeMonthItems As Object, _
    ByVal sixMonthItems As Object)
    
    On Error GoTo ErrHandler
    
    Dim key As Variant
    Dim qcNumber As String
    Dim qcStatus As String
    Dim nilStockFlag As String
    Dim lastRetestDateStr As String
    Dim lastRetestDate As Date
    Dim daysUntilRetest As Long
    Dim rmCode As String
    Dim standardName As String
    Dim qcBalance As String
    Dim masterUoM As String
    Dim isEligible As Boolean
    Dim itemData As Object
    Dim todayDate As Date
    Dim lastRow As Long
    Dim r As Long
    Dim colQCNumber As Long
    Dim colQCStatus As Long
    Dim colNilStockFlag As Long
    Dim colLastRetestDate As Long
    Dim colRMCode As Long
    Dim colStandardName As Long
    Dim colQCBalance As Long
    Dim colMasterUoM As Long
    Dim uniqueKey As String

    todayDate = Date
    lastRow = wsQCData.Cells(wsQCData.Rows.count, 1).End(xlUp).row
    
    colQCNumber = FindHeaderCol(wsQCData, 1, HEADER_QC_NUMBER)
    colQCStatus = FindHeaderCol(wsQCData, 1, HEADER_QC_STATUS_FLAG)
    colNilStockFlag = FindHeaderCol(wsQCData, 1, HEADER_NIL_STOCK_FLAG)
    colLastRetestDate = FindHeaderCol(wsQCData, 1, HEADER_LAST_RETEST_DATE)
    colRMCode = FindHeaderCol(wsQCData, 1, HEADER_RM_CODE)
    colStandardName = FindHeaderCol(wsQCData, 1, HEADER_STANDARD_NAME)
    colQCBalance = FindHeaderCol(wsQCData, 1, HEADER_QC_BALANCE)
    colMasterUoM = FindHeaderCol(wsQCData, 1, HEADER_MASTER_UOM)
    
    If colQCNumber = 0 Or colQCStatus = 0 Or colNilStockFlag = 0 Or colLastRetestDate = 0 Then
        Err.Raise vbObjectError + 106, , "Required headers not found in QC_Number worksheet"
    End If
    
    For r = 2 To lastRow
        qcStatus = UCase(Trim(Replace(wsQCData.Cells(r, colQCStatus).value, """", "")))
        nilStockFlag = UCase(Trim(Replace(wsQCData.Cells(r, colNilStockFlag).value, """", "")))
        lastRetestDateStr = Trim(Replace(wsQCData.Cells(r, colLastRetestDate).value, """", ""))
        
        isEligible = False
        
        If qcStatus <> "REJECTED" And qcStatus <> "CANCELLED" Then
            If nilStockFlag <> "Y" Then
                isEligible = True
            End If
        End If
        
        If Not isEligible Then GoTo NextRecord
        If lastRetestDateStr = "" Then GoTo NextRecord
        
        On Error Resume Next
        lastRetestDate = CDate(lastRetestDateStr)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextRecord
        End If
        On Error GoTo ErrHandler
        
        daysUntilRetest = DateDiff("d", todayDate, lastRetestDate)
        
        qcNumber = Trim(Replace(wsQCData.Cells(r, colQCNumber).value, """", ""))
        rmCode = Trim(Replace(wsQCData.Cells(r, colRMCode).value, """", ""))
        standardName = Trim(Replace(wsQCData.Cells(r, colStandardName).value, """", ""))
        qcBalance = Trim(Replace(wsQCData.Cells(r, colQCBalance).value, """", ""))
        masterUoM = Trim(Replace(wsQCData.Cells(r, colMasterUoM).value, """", ""))
        
        uniqueKey = qcNumber & "_" & r
        
        Set itemData = CreateObject("Scripting.Dictionary")
        itemData("Code") = rmCode
        itemData("StandardName") = standardName
        itemData("QCNumber") = qcNumber
        itemData("QCBalance") = qcBalance
        itemData("MasterUoM") = masterUoM
        itemData("RetestDate") = lastRetestDateStr
        itemData("DaysUntilRetest") = daysUntilRetest
        itemData("QCStatus") = qcStatus
        
        If daysUntilRetest < 0 Then
            pastDueItems.Add uniqueKey, itemData
        ElseIf daysUntilRetest <= 30 Then
            oneMonthItems.Add uniqueKey, itemData
        ElseIf daysUntilRetest <= 90 Then
            threeMonthItems.Add uniqueKey, itemData
        ElseIf daysUntilRetest <= 180 Then
            sixMonthItems.Add uniqueKey, itemData
        End If
        
NextRecord:
    Next r
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Set itemData = Nothing
    Exit Sub
End Sub

Private Sub OutputRetestReportData( _
    ByVal ws As Worksheet, _
    ByVal bucketName As String, _
    ByVal items As Object, _
    ByRef outputRow As Long)
    
    On Error GoTo ErrHandler
    
    Dim key As Variant
    Dim itemData As Object
    Dim sortedKeys() As String
    Dim i As Long
    Dim count As Long
    Dim paddedCode As String
    
    If items.count = 0 Then GoTo CleanExit
    
    ReDim sortedKeys(0 To items.count - 1)
    count = 0
    
    For Each key In items.keys
        sortedKeys(count) = CStr(key)
        count = count + 1
    Next key
    
    QuickSortItemsByCode items, sortedKeys, 0, UBound(sortedKeys)
    
    For i = 0 To UBound(sortedKeys)
        Set itemData = items(sortedKeys(i))
        paddedCode = RM_Pad4(itemData("Code"))
        
        ws.Cells(outputRow, 8).value = bucketName
        ws.Cells(outputRow, 9).NumberFormat = "@"
        ws.Cells(outputRow, 9).value = paddedCode
        ws.Cells(outputRow, 10).value = itemData("StandardName")
        ws.Cells(outputRow, 11).value = itemData("QCNumber")
        ws.Cells(outputRow, 12).value = itemData("QCBalance")
        ws.Cells(outputRow, 13).value = itemData("MasterUoM")
        ws.Cells(outputRow, 14).HorizontalAlignment = xlRight
        ws.Cells(outputRow, 14).value = itemData("RetestDate")
        ws.Cells(outputRow, 15).value = itemData("QCStatus")
        
        outputRow = outputRow + 1
    Next i
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Set itemData = Nothing
    Exit Sub
End Sub

Private Sub QuickSortItemsByCode( _
    ByVal items As Object, _
    ByRef keys() As String, _
    ByVal low As Long, _
    ByVal high As Long)
    
    On Error GoTo ErrHandler
    
    Dim pivot As String
    Dim i As Long
    Dim j As Long
    Dim temp As String
    Dim pivotCode As String
    
    If low < high Then
        pivot = keys((low + high) \ 2)
        pivotCode = items(pivot)("Code")
        i = low
        j = high
        
        Do While i <= j
            Do While items(keys(i))("Code") < pivotCode
                i = i + 1
            Loop
            
            Do While items(keys(j))("Code") > pivotCode
                j = j - 1
            Loop
            
            If i <= j Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If low < j Then QuickSortItemsByCode items, keys, low, j
        If i < high Then QuickSortItemsByCode items, keys, i, high
    End If
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Exit Sub
End Sub

Public Sub GenerateManufacturerExpiryReport()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim wsQCData As Worksheet
    Dim wsRMMaster As Worksheet
    Dim pastDueItems As Object
    Dim oneMonthItems As Object
    Dim threeMonthItems As Object
    Dim sixMonthItems As Object
    Dim outputRow As Long
    
    LoadQCNumberData
    
    Set ws = ActiveSheet
    Set wsQCData = ThisWorkbook.Worksheets("QC_Number")
    Set wsRMMaster = ThisWorkbook.Worksheets("RM_Master")
    Set pastDueItems = CreateObject("Scripting.Dictionary")
    Set oneMonthItems = CreateObject("Scripting.Dictionary")
    Set threeMonthItems = CreateObject("Scripting.Dictionary")
    Set sixMonthItems = CreateObject("Scripting.Dictionary")
    
    CategorizeManufacturerExpiryItems wsQCData, wsRMMaster, pastDueItems, oneMonthItems, threeMonthItems, sixMonthItems
    
    UnprotectActiveSheet
    
    ws.Range("C6").value = pastDueItems.count
    ws.Range("D6").value = oneMonthItems.count
    ws.Range("E6").value = threeMonthItems.count
    ws.Range("F6").value = sixMonthItems.count
    
    ws.Range("H10:N10000").ClearContents
    
    outputRow = 10
    OutputManufacturerExpiryReportData ws, "Past Due", pastDueItems, outputRow
    OutputManufacturerExpiryReportData ws, "1 Month", oneMonthItems, outputRow
    OutputManufacturerExpiryReportData ws, "3 Month", threeMonthItems, outputRow
    OutputManufacturerExpiryReportData ws, "6 Month", sixMonthItems, outputRow
    
    ProtectActiveSheet
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateManufacturerExpiryReport"
    GoTo CleanExit

CleanExit:
    Set ws = Nothing
    Set wsQCData = Nothing
    Set wsRMMaster = Nothing
    Set pastDueItems = Nothing
    Set oneMonthItems = Nothing
    Set threeMonthItems = Nothing
    Set sixMonthItems = Nothing
    Exit Sub
End Sub

Private Sub CategorizeManufacturerExpiryItems( _
    ByVal wsQCData As Worksheet, _
    ByVal wsRMMaster As Worksheet, _
    ByVal pastDueItems As Object, _
    ByVal oneMonthItems As Object, _
    ByVal threeMonthItems As Object, _
    ByVal sixMonthItems As Object)
    
    On Error GoTo ErrHandler
    
    Dim rmCode As String
    Dim qcNumberRequired As String
    Dim testingFlag As String
    Dim qcStatus As String
    Dim mfgExpiryStr As String
    Dim mfgExpiryDate As Date
    Dim daysUntilExpiry As Long
    Dim qcNumber As String
    Dim standardName As String
    Dim qcBalance As String
    Dim masterUoM As String
    Dim itemData As Object
    Dim todayDate As Date
    Dim lastRow As Long
    Dim r As Long
    Dim rmRow As Long
    Dim rmLastRow As Long
    Dim uniqueKey As String
    Dim colRMCode As Long
    Dim colMfgExpiry As Long
    Dim colQCNumber As Long
    Dim colQCStatus As Long
    Dim colStandardName As Long
    Dim colQCBalance As Long
    Dim colMasterUoM As Long
    Dim colTestingFlag As Long
    Dim colQCNumberRequired As Long
    Dim colRMCodeMaster As Long
    Dim rmFlagDict As Object
    
    todayDate = Date
    
    Set rmFlagDict = CreateObject("Scripting.Dictionary")
    
    colRMCode = FindHeaderCol(wsQCData, 1, HEADER_RM_CODE)
    colQCNumber = FindHeaderCol(wsQCData, 1, HEADER_QC_NUMBER)
    colStandardName = FindHeaderCol(wsQCData, 1, HEADER_STANDARD_NAME)
    colQCBalance = FindHeaderCol(wsQCData, 1, HEADER_QC_BALANCE)
    colMasterUoM = FindHeaderCol(wsQCData, 1, HEADER_MASTER_UOM)
    colMfgExpiry = FindHeaderCol(wsQCData, 1, HEADER_MANUFACTURER_EXPIRY)
    colQCStatus = FindHeaderCol(wsQCData, 1, HEADER_QC_STATUS_FLAG)
    
    If colRMCode = 0 Or colMfgExpiry = 0 Or colQCNumber = 0 Or colQCStatus = 0 Then
        Err.Raise vbObjectError + 106, , "Required headers not found in QC_Number worksheet"
    End If
    
    colTestingFlag = FindHeaderCol(wsRMMaster, 1, RMM_HEADER_TESTING_FLAG)
    colQCNumberRequired = FindHeaderCol(wsRMMaster, 1, RMM_HEADER_QC_NUMBER_REQUIRED)
    colRMCodeMaster = FindHeaderCol(wsRMMaster, 1, RMM_HEADER_RM_CODE)
    
    If colTestingFlag = 0 Or colQCNumberRequired = 0 Or colRMCodeMaster = 0 Then
        Err.Raise vbObjectError + 107, , "Required headers not found in RM_Master worksheet"
    End If
    
    rmLastRow = wsRMMaster.Cells(wsRMMaster.Rows.count, 1).End(xlUp).row
    For rmRow = 2 To rmLastRow
        rmCode = Trim(Replace(wsRMMaster.Cells(rmRow, colRMCodeMaster).value, """", ""))
        qcNumberRequired = UCase(Trim(Replace(wsRMMaster.Cells(rmRow, colQCNumberRequired).value, """", "")))
        testingFlag = UCase(Trim(Replace(wsRMMaster.Cells(rmRow, colTestingFlag).value, """", "")))
        
        If qcNumberRequired = "Y" And testingFlag = "N" Then
            rmFlagDict(rmCode) = True
        End If
    Next rmRow
    
    lastRow = wsQCData.Cells(wsQCData.Rows.count, 1).End(xlUp).row
    
    For r = 2 To lastRow
        rmCode = Trim(Replace(wsQCData.Cells(r, colRMCode).value, """", ""))
        mfgExpiryStr = Trim(Replace(wsQCData.Cells(r, colMfgExpiry).value, """", ""))
        qcStatus = UCase(Trim(Replace(wsQCData.Cells(r, colQCStatus).value, """", "")))
        
        If Not rmFlagDict.Exists(rmCode) Then GoTo NextRecord
        If mfgExpiryStr = "" Then GoTo NextRecord
        
        On Error Resume Next
        mfgExpiryDate = CDate(mfgExpiryStr)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextRecord
        End If
        On Error GoTo ErrHandler
        
        daysUntilExpiry = DateDiff("d", todayDate, mfgExpiryDate)
        
        qcNumber = Trim(Replace(wsQCData.Cells(r, colQCNumber).value, """", ""))
        standardName = Trim(Replace(wsQCData.Cells(r, colStandardName).value, """", ""))
        qcBalance = Trim(Replace(wsQCData.Cells(r, colQCBalance).value, """", ""))
        masterUoM = Trim(Replace(wsQCData.Cells(r, colMasterUoM).value, """", ""))
        
        uniqueKey = qcNumber & "_" & r
        
        Set itemData = CreateObject("Scripting.Dictionary")
        itemData("Code") = rmCode
        itemData("StandardName") = standardName
        itemData("QCNumber") = qcNumber
        itemData("QCBalance") = qcBalance
        itemData("MasterUoM") = masterUoM
        itemData("ExpiryDate") = mfgExpiryStr
        itemData("DaysUntilExpiry") = daysUntilExpiry
        itemData("QCStatus") = qcStatus
        
        If daysUntilExpiry < 0 Then
            pastDueItems.Add uniqueKey, itemData
        ElseIf daysUntilExpiry <= 30 Then
            oneMonthItems.Add uniqueKey, itemData
        ElseIf daysUntilExpiry <= 90 Then
            threeMonthItems.Add uniqueKey, itemData
        ElseIf daysUntilExpiry <= 180 Then
            sixMonthItems.Add uniqueKey, itemData
        End If
        
NextRecord:
    Next r
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Set itemData = Nothing
    Set rmFlagDict = Nothing
    Exit Sub
End Sub

Private Sub OutputManufacturerExpiryReportData( _
    ByVal ws As Worksheet, _
    ByVal bucketName As String, _
    ByVal items As Object, _
    ByRef outputRow As Long)
    
    On Error GoTo ErrHandler
    
    Dim key As Variant
    Dim itemData As Object
    Dim sortedKeys() As String
    Dim i As Long
    Dim count As Long
    Dim paddedCode As String
    
    If items.count = 0 Then GoTo CleanExit
    
    ReDim sortedKeys(0 To items.count - 1)
    count = 0
    
    For Each key In items.keys
        sortedKeys(count) = CStr(key)
        count = count + 1
    Next key
    
    QuickSortItemsByCode items, sortedKeys, 0, UBound(sortedKeys)
    
    For i = 0 To UBound(sortedKeys)
        Set itemData = items(sortedKeys(i))
        paddedCode = RM_Pad4(itemData("Code"))
        
        ws.Cells(outputRow, 8).value = bucketName
        ws.Cells(outputRow, 9).NumberFormat = "@"
        ws.Cells(outputRow, 9).value = paddedCode
        ws.Cells(outputRow, 10).value = itemData("StandardName")
        ws.Cells(outputRow, 11).value = itemData("QCNumber")
        ws.Cells(outputRow, 12).value = itemData("QCBalance")
        ws.Cells(outputRow, 13).value = itemData("MasterUoM")
        ws.Cells(outputRow, 14).HorizontalAlignment = xlRight
        ws.Cells(outputRow, 14).value = itemData("ExpiryDate")
        ws.Cells(outputRow, 15).value = itemData("QCStatus")
        
        outputRow = outputRow + 1
    Next i
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Set itemData = Nothing
    Exit Sub
End Sub

