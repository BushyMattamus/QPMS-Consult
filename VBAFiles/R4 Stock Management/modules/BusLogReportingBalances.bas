Attribute VB_Name = "BusLogReportingBalances"
' ========================================================================
' BUSINESS LOGISTICS REPORTING
' ========================================================================
' Generates stock movement reports from cached QC Number and RM Trans data
' Displays current and previous month activity with stock reconciliation
' ========================================================================

Option Explicit

Public Sub GenerateBalancesReport()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim conn As Object
    Dim rs As Object
    Dim fso As Object
    Dim tempQCPath As String
    Dim tempRMTransPath As String
    Dim tempRMMasterPath As String
    Dim tempFolderPath As String
    Dim currentYear As Long
    Dim currentMonth As Long
    Dim prevYear As Long
    Dim prevMonth As Long
    Dim currentMonthEnd As Date
    Dim prevMonthEnd As Date
    Dim activeQCNumbers As Object
    Dim rmCodeData As Object
    Dim outputRow As Long
    Dim yearStr As String
    Dim monthStr As String
    Dim sourceFolderPath As String
    
    Set ws = ActiveSheet
    Set activeQCNumbers = CreateObject("Scripting.Dictionary")
    Set rmCodeData = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
        
    yearStr = GetFieldValByLabelC("Year:")
    monthStr = GetFieldValByLabelC("Month:")
    
    Debug.Print "DEBUG 1: Year=" & yearStr & ", Month=" & monthStr
    
    If Not IsNumeric(yearStr) Or Not IsNumeric(monthStr) Then
        MsgBox "Invalid Year or Month value", vbExclamation, "GenerateBalancesReport"
        GoTo CleanExit
    End If
    
    currentYear = CLng(yearStr)
    currentMonth = CLng(monthStr)
    
    currentMonthEnd = DateSerial(currentYear, currentMonth + 1, 0)
    
    If currentMonth = 1 Then
        prevYear = currentYear - 1
        prevMonth = 12
    Else
        prevYear = currentYear
        prevMonth = currentMonth - 1
    End If
    
    prevMonthEnd = DateSerial(prevYear, prevMonth + 1, 0)
    
    Debug.Print "DEBUG 2: currentMonthEnd=" & currentMonthEnd & ", prevMonthEnd=" & prevMonthEnd

    tempQCPath = CreateTemporaryFileCopy(FILE_QC_NUMBER)
    Debug.Print "DEBUG 3: tempQCPath=" & tempQCPath
    If tempQCPath = "" Then
        MsgBox "Failed to create temporary copy of QC_Number.csv", vbCritical, "GenerateBalancesReport"
        GoTo CleanExit
    End If
    
    tempRMTransPath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    Debug.Print "DEBUG 4: tempRMTransPath=" & tempRMTransPath
    If tempRMTransPath = "" Then
        MsgBox "Failed to create temporary copy of RM_Trans.csv", vbCritical, "GenerateBalancesReport"
        GoTo CleanExit
    End If
    
    sourceFolderPath = MASTER_DATA_PATH & FOLDER_TABLES
    Debug.Print "DEBUG 5: sourceFolderPath=" & sourceFolderPath
    tempRMMasterPath = CreateTemporaryFileCopy(FILE_RM_MASTER, sourceFolderPath)
    Debug.Print "DEBUG 6: tempRMMasterPath=" & tempRMMasterPath
    If tempRMMasterPath = "" Then
        MsgBox "Failed to create temporary copy of Raw_Material_Master.csv", vbCritical, "GenerateBalancesReport"
        GoTo CleanExit
    End If
    
    tempFolderPath = fso.GetParentFolderName(tempQCPath)
    Debug.Print "DEBUG 7: tempFolderPath=" & tempFolderPath
    
    Set conn = OpenADOConnection(tempFolderPath)
    If conn Is Nothing Then
        MsgBox "Failed to open ADO connection", vbCritical, "GenerateBalancesReport"
        GoTo CleanExit
    End If
    Debug.Print "DEBUG 8: ADO connection opened successfully"
    
    Debug.Print "DEBUG 9: QC filename=" & fso.GetFileName(tempQCPath)
    IdentifyActiveQCNumbersFromADO conn, activeQCNumbers, currentMonthEnd, prevMonthEnd, _
                                    fso.GetFileName(tempQCPath)
    Debug.Print "DEBUG 10: activeQCNumbers.Count=" & activeQCNumbers.count
    
    If activeQCNumbers.count = 0 Then
        MsgBox "No active QC numbers found for the selected period", vbInformation, "GenerateBalancesReport"
        GoTo CleanExit
    End If
    
    AggregateRMCodeDataFromADO conn, activeQCNumbers, rmCodeData, currentMonthEnd, prevMonthEnd, _
                               fso.GetFileName(tempQCPath), fso.GetFileName(tempRMTransPath), _
                               fso.GetFileName(tempRMMasterPath)
    Debug.Print "DEBUG 11: rmCodeData.Count=" & rmCodeData.count
        
    ws.Range("E8:L10000").ClearContents
    ws.Range("E7:L10000").Locked = False
    
    outputRow = 8
    OutputReportData ws, rmCodeData, outputRow
    Debug.Print "DEBUG 12: Output complete, final row=" & outputRow
        
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateBalancesReport"
    Debug.Print "DEBUG ERROR: " & Err.Number & " - " & Err.Description
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If Not conn.State = 0 Then conn.Close
    End If
    If Not fso Is Nothing Then
        If tempQCPath <> "" And fso.FileExists(tempQCPath) Then fso.DeleteFile tempQCPath
        If tempRMTransPath <> "" And fso.FileExists(tempRMTransPath) Then fso.DeleteFile tempRMTransPath
        If tempRMMasterPath <> "" And fso.FileExists(tempRMMasterPath) Then fso.DeleteFile tempRMMasterPath
    End If
    Set rs = Nothing
    Set conn = Nothing
    Set fso = Nothing
    Set ws = Nothing
    Set activeQCNumbers = Nothing
    Set rmCodeData = Nothing
    Exit Sub
End Sub

Private Sub IdentifyActiveQCNumbersFromADO( _
    ByVal conn As Object, _
    ByVal activeQCNumbers As Object, _
    ByVal currentMonthEnd As Date, _
    ByVal prevMonthEnd As Date, _
    ByVal qcFileName As String)
    
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim sql As String
    Dim qcNumber As String
    Dim grDate As Date
    Dim nilStockFlag As String
    Dim nilStockDate As Date
    Dim isActive As Boolean
    Dim recordCount As Long
    Dim hasBOM As Boolean
    Dim qcNumHeader As String
    
    hasBOM = False
    qcNumHeader = HEADER_QC_NUMBER
    
    sql = "SELECT [" & qcNumHeader & "], " & _
          "[" & HEADER_GR_DATE & "], " & _
          "[" & HEADER_NIL_STOCK_FLAG & "], " & _
          "[" & HEADER_NIL_STOCK_FLAG_DATE & "] " & _
          "FROM [" & qcFileName & "] " & _
          "WHERE [" & HEADER_GR_DATE & "] IS NOT NULL AND [" & HEADER_GR_DATE & "] <> ''"
    
    On Error Resume Next
    Set rs = conn.Execute(sql)
    
    If Err.Number <> 0 Then
        Err.Clear
        hasBOM = True
        qcNumHeader = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) & HEADER_QC_NUMBER
        
        sql = "SELECT [" & qcNumHeader & "], " & _
              "[" & HEADER_GR_DATE & "], " & _
              "[" & HEADER_NIL_STOCK_FLAG & "], " & _
              "[" & HEADER_NIL_STOCK_FLAG_DATE & "] " & _
              "FROM [" & qcFileName & "] " & _
              "WHERE [" & HEADER_GR_DATE & "] IS NOT NULL AND [" & HEADER_GR_DATE & "] <> ''"
        
        Set rs = conn.Execute(sql)
        If Err.Number <> 0 Then
            Debug.Print "DEBUG IdentifyActive ERROR: Unable to query file - " & Err.Number & " - " & Err.Description
            Err.Clear
            GoTo CleanExit
        End If
    End If
    On Error GoTo ErrHandler
    
    Debug.Print "DEBUG IdentifyActive: BOM detected=" & hasBOM
    Debug.Print "DEBUG IdentifyActive: rs.EOF=" & rs.EOF
    
    recordCount = 0
    Do While Not rs.EOF
        recordCount = recordCount + 1
        
        On Error Resume Next
        grDate = CDate(rs.Fields(HEADER_GR_DATE).value)
        If Err.Number <> 0 Then
            Debug.Print "DEBUG: Failed to parse GR_Date: " & rs.Fields(HEADER_GR_DATE).value
            Err.Clear
            GoTo NextRecord
        End If
        On Error GoTo ErrHandler
        
        If grDate > currentMonthEnd Then GoTo NextRecord
        
        qcNumber = Trim(CStr(rs.Fields(qcNumHeader).value))
        nilStockFlag = UCase(Trim(CStr(rs.Fields(HEADER_NIL_STOCK_FLAG).value & "")))
        
        isActive = False
        
        If nilStockFlag = "" Or nilStockFlag = "N" Then
            isActive = True
        ElseIf nilStockFlag = "Y" Then
            On Error Resume Next
            nilStockDate = CDate(rs.Fields(HEADER_NIL_STOCK_FLAG_DATE).value)
            If Err.Number = 0 Then
                If nilStockDate >= prevMonthEnd Then
                    isActive = True
                End If
            End If
            Err.Clear
            On Error GoTo ErrHandler
        End If
        
        If isActive Then
            activeQCNumbers(qcNumber) = True
        End If
        
NextRecord:
        rs.MoveNext
    Loop
    
    Debug.Print "DEBUG IdentifyActive: recordCount=" & recordCount & ", activeQCNumbers.Count=" & activeQCNumbers.count
    
    GoTo CleanExit

ErrHandler:
    Debug.Print "DEBUG IdentifyActive ERROR: " & Err.Number & " - " & Err.Description
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
    End If
    Set rs = Nothing
    Exit Sub
End Sub

Private Sub AggregateRMCodeDataFromADO( _
    ByVal conn As Object, _
    ByVal activeQCNumbers As Object, _
    ByVal rmCodeData As Object, _
    ByVal currentMonthEnd As Date, _
    ByVal prevMonthEnd As Date, _
    ByVal qcFileName As String, _
    ByVal rmTransFileName As String, _
    ByVal rmMasterFileName As String)
    
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim sql As String
    Dim qcNumber As String
    Dim rmCode As String
    Dim stdName As String
    Dim masterUoM As String
    Dim category As String
    Dim transDate As Date
    Dim quantity As Double
    Dim rmData As Object
    Dim isCurrentMonth As Boolean
    Dim transDateStr As String
    Dim hasBOMQC As Boolean
    Dim hasBOMRM As Boolean
    Dim qcNumHeaderQC As String
    Dim qcNumHeaderRM As String
    
    hasBOMQC = False
    qcNumHeaderQC = HEADER_QC_NUMBER
    
    sql = "SELECT [" & qcNumHeaderQC & "], " & _
          "[" & HEADER_RM_CODE & "], " & _
          "[" & HEADER_STANDARD_NAME & "], " & _
          "[" & HEADER_MASTER_UOM & "] " & _
          "FROM [" & qcFileName & "]"
    
    On Error Resume Next
    Set rs = conn.Execute(sql)
    
    If Err.Number <> 0 Then
        Err.Clear
        hasBOMQC = True
        qcNumHeaderQC = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) & HEADER_QC_NUMBER
        
        sql = "SELECT [" & qcNumHeaderQC & "], " & _
              "[" & HEADER_RM_CODE & "], " & _
              "[" & HEADER_STANDARD_NAME & "], " & _
              "[" & HEADER_MASTER_UOM & "] " & _
              "FROM [" & qcFileName & "]"
        
        Set rs = conn.Execute(sql)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo CleanExit
        End If
    End If
    On Error GoTo ErrHandler
    
    Do While Not rs.EOF
        qcNumber = Trim(CStr(rs.Fields(qcNumHeaderQC).value))
        
        If activeQCNumbers.Exists(qcNumber) Then
            rmCode = Trim(CStr(rs.Fields(HEADER_RM_CODE).value))
            
            If Not rmCodeData.Exists(rmCode) Then
                stdName = Trim(CStr(rs.Fields(HEADER_STANDARD_NAME).value))
                masterUoM = Trim(CStr(rs.Fields(HEADER_MASTER_UOM).value))
                category = GetRMCategoryFromADO(conn, rmCode, rmMasterFileName)
                
                Set rmData = CreateObject("Scripting.Dictionary")
                rmData("StandardName") = stdName
                rmData("MasterUoM") = masterUoM
                rmData("Category") = category
                rmData("CurrentIns") = 0#
                rmData("CurrentOuts") = 0#
                rmData("PrevIns") = 0#
                rmData("PrevOuts") = 0#
                Set rmData("QCNumbers") = CreateObject("Scripting.Dictionary")
                rmCodeData.Add rmCode, rmData
            Else
                Set rmData = rmCodeData(rmCode)
            End If
            
            rmData("QCNumbers")(qcNumber) = True
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    hasBOMRM = False
    qcNumHeaderRM = RMT_HEADER_QC_NUMBER
    
    sql = "SELECT [" & qcNumHeaderRM & "], " & _
          "[" & RMT_HEADER_QUANTITY & "], " & _
          "[" & RMT_HEADER_TRANS_DT & "] " & _
          "FROM [" & rmTransFileName & "] " & _
          "WHERE [" & qcNumHeaderRM & "] IS NOT NULL AND [" & qcNumHeaderRM & "] <> ''"
    
    On Error Resume Next
    Set rs = conn.Execute(sql)
    
    If Err.Number <> 0 Then
        Err.Clear
        hasBOMRM = True
        qcNumHeaderRM = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) & RMT_HEADER_QC_NUMBER
        
        sql = "SELECT [" & qcNumHeaderRM & "], " & _
              "[" & RMT_HEADER_QUANTITY & "], " & _
              "[" & RMT_HEADER_TRANS_DT & "] " & _
              "FROM [" & rmTransFileName & "] " & _
              "WHERE [" & qcNumHeaderRM & "] IS NOT NULL AND [" & qcNumHeaderRM & "] <> ''"
        
        Set rs = conn.Execute(sql)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo CleanExit
        End If
    End If
    On Error GoTo ErrHandler
    
    Do While Not rs.EOF
        qcNumber = Trim(CStr(rs.Fields(qcNumHeaderRM).value))
        
        If activeQCNumbers.Exists(qcNumber) Then
            rmCode = FindRMCodeForQCNumber(rmCodeData, qcNumber)
            
            If rmCode <> "" Then
                transDateStr = Trim(CStr(rs.Fields(RMT_HEADER_TRANS_DT).value))
                
                If Len(transDateStr) >= 10 Then
                    transDateStr = Left$(transDateStr, 10)
                    
                    On Error Resume Next
                    transDate = CDate(transDateStr)
                    If Err.Number = 0 Then
                        If transDate <= currentMonthEnd Then
                            Set rmData = rmCodeData(rmCode)
                            
                            quantity = CDbl(rs.Fields(RMT_HEADER_QUANTITY).value)
                            If Err.Number <> 0 Then quantity = 0#
                            
                            isCurrentMonth = (transDate > prevMonthEnd)
                            
                            If quantity > 0# Then
                                If isCurrentMonth Then
                                    rmData("CurrentIns") = rmData("CurrentIns") + quantity
                                Else
                                    rmData("PrevIns") = rmData("PrevIns") + quantity
                                End If
                            Else
                                If isCurrentMonth Then
                                    rmData("CurrentOuts") = rmData("CurrentOuts") + quantity
                                Else
                                    rmData("PrevOuts") = rmData("PrevOuts") + quantity
                                End If
                            End If
                        End If
                    End If
                    Err.Clear
                    On Error GoTo ErrHandler
                End If
            End If
        End If
        
        rs.MoveNext
    Loop
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
    End If
    Set rs = Nothing
    Set rmData = Nothing
    Exit Sub
End Sub

Private Function FindRMCodeForQCNumber(ByVal rmCodeData As Object, ByVal qcNumber As String) As String
    On Error GoTo ErrHandler
    
    Dim rmCodeKey As Variant
    Dim rmData As Object
    
    For Each rmCodeKey In rmCodeData.keys
        Set rmData = rmCodeData(rmCodeKey)
        If rmData("QCNumbers").Exists(qcNumber) Then
            FindRMCodeForQCNumber = CStr(rmCodeKey)
            GoTo CleanExit
        End If
    Next rmCodeKey
    
    FindRMCodeForQCNumber = ""
    GoTo CleanExit

ErrHandler:
    FindRMCodeForQCNumber = ""
    GoTo CleanExit

CleanExit:
    Set rmData = Nothing
    Exit Function
End Function

Private Function GetRMCategoryFromADO(ByVal conn As Object, ByVal rmCode As String, ByVal rmMasterFileName As String) As String
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim sql As String
    Dim searchCode As String
    Dim hasBOM As Boolean
    Dim rmCodeHeader As String
    
    searchCode = RM_Pad4(Trim$(rmCode))
    
    hasBOM = False
    rmCodeHeader = RMM_HEADER_RM_CODE
    
    sql = "SELECT [" & RMM_HEADER_CATEGORY & "] " & _
          "FROM [" & rmMasterFileName & "] " & _
          "WHERE [" & rmCodeHeader & "] = '" & Replace(searchCode, "'", "''") & "'"
    
    On Error Resume Next
    Set rs = conn.Execute(sql)
    
    If Err.Number <> 0 Then
        Err.Clear
        hasBOM = True
        rmCodeHeader = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) & RMM_HEADER_RM_CODE
        
        sql = "SELECT [" & RMM_HEADER_CATEGORY & "] " & _
              "FROM [" & rmMasterFileName & "] " & _
              "WHERE [" & rmCodeHeader & "] = '" & Replace(searchCode, "'", "''") & "'"
        
        Set rs = conn.Execute(sql)
        If Err.Number <> 0 Then
            Err.Clear
            GetRMCategoryFromADO = ""
            GoTo CleanExit
        End If
    End If
    On Error GoTo ErrHandler
    
    If Not rs.EOF Then
        GetRMCategoryFromADO = Trim(CStr(rs.Fields(RMM_HEADER_CATEGORY).value))
    Else
        GetRMCategoryFromADO = ""
    End If
    
    GoTo CleanExit

ErrHandler:
    GetRMCategoryFromADO = ""
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
    End If
    Set rs = Nothing
    Exit Function
End Function

Private Sub OutputReportData( _
    ByVal ws As Worksheet, _
    ByVal rmCodeData As Object, _
    ByRef outputRow As Long)
    
    On Error GoTo ErrHandler
    
    Dim rmCode As Variant
    Dim rmData As Object
    Dim prevSOH As Double
    Dim currentSOH As Double
    Dim netChange As Double
    Dim sortedCodes() As String
    Dim i As Long
    Dim count As Long
    
    If rmCodeData.count = 0 Then GoTo CleanExit
    
    ReDim sortedCodes(0 To rmCodeData.count - 1)
    count = 0
    
    For Each rmCode In rmCodeData.keys
        sortedCodes(count) = CStr(rmCode)
        count = count + 1
    Next rmCode
    
    QuickSortStrings sortedCodes, 0, UBound(sortedCodes)
    
    For i = 0 To UBound(sortedCodes)
        Set rmData = rmCodeData(sortedCodes(i))
        
        prevSOH = rmData("PrevIns") + rmData("PrevOuts")
        currentSOH = prevSOH + rmData("CurrentIns") + rmData("CurrentOuts")
        netChange = currentSOH - prevSOH
        
        ws.Cells(outputRow, 5).value = rmData("Category")
        ws.Cells(outputRow, 6).NumberFormat = "@"
        ws.Cells(outputRow, 6).value = RM_Pad4(sortedCodes(i))
        ws.Cells(outputRow, 7).value = rmData("StandardName")
        ws.Cells(outputRow, 8).value = Round(currentSOH, 3)
        ws.Cells(outputRow, 9).value = Round(rmData("CurrentIns"), 3)
        ws.Cells(outputRow, 10).value = Round(rmData("CurrentOuts"), 3)
        ws.Cells(outputRow, 11).value = Round(prevSOH, 3)
        ws.Cells(outputRow, 12).value = Round(netChange, 3)
        ws.Cells(outputRow, 13).value = rmData("MasterUoM")
        
        outputRow = outputRow + 1
    Next i
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Set rmData = Nothing
    Exit Sub
End Sub

Private Sub QuickSortStrings(ByRef arr() As String, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrHandler
    
    Dim pivot As String
    Dim i As Long
    Dim j As Long
    Dim temp As String
    
    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high
        
        Do While i <= j
            Do While arr(i) < pivot
                i = i + 1
            Loop
            
            Do While arr(j) > pivot
                j = j - 1
            Loop
            
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If low < j Then QuickSortStrings arr, low, j
        If i < high Then QuickSortStrings arr, i, high
    End If
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Exit Sub
End Sub

Public Sub GenerateCurrentSOHReport()
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim ws As Worksheet
    Dim qcCache As Object
    Dim rmCache As Object
    Dim rmCodeData As Object
    Dim rmCode As Variant
    Dim rmData As Object
    Dim sortedCodes() As String
    Dim i As Long
    Dim count As Long
    Dim outputRow As Long
    Dim currentSOH As Double
    
    ' --- Object initialization ---
    Set ws = ActiveSheet
    Set rmCodeData = CreateObject("Scripting.Dictionary")
    
    ' --- Refresh and load caches ---
    RefreshQCNumberCache
    RefreshRMTransCache
    
    Set qcCache = GetQCNumberCache()
    Set rmCache = GetRMTransCache()
    
    If qcCache Is Nothing Or rmCache Is Nothing Then
        MsgBox "Failed to load cached data", vbCritical, "GenerateCurrentSOHReport"
        GoTo CleanExit
    End If
    
    ' --- Aggregate all SOH data ---
    AggregateAllSOHData qcCache, rmCache, rmCodeData
    
    ' --- Output to worksheet ---
    ws.Range("E4:I10000").ClearContents
    ws.Range("E3:I10000").Locked = False
    
    If rmCodeData.count = 0 Then
        GoTo CleanExit
    End If
    
    ReDim sortedCodes(0 To rmCodeData.count - 1)
    count = 0
    
    For Each rmCode In rmCodeData.keys
        sortedCodes(count) = CStr(rmCode)
        count = count + 1
    Next rmCode
    
    QuickSortStrings sortedCodes, 0, UBound(sortedCodes)
    
    outputRow = 4
    For i = 0 To UBound(sortedCodes)
        Set rmData = rmCodeData(sortedCodes(i))
        
        currentSOH = rmData("TotalIns") + rmData("TotalOuts")
        
        ws.Cells(outputRow, 5).value = rmData("Category")
        ws.Cells(outputRow, 6).NumberFormat = "@"
        ws.Cells(outputRow, 6).value = RM_Pad4(sortedCodes(i))
        ws.Cells(outputRow, 7).value = rmData("StandardName")
        ws.Cells(outputRow, 8).value = Round(currentSOH, 3)
        ws.Cells(outputRow, 9).value = rmData("MasterUoM")
        
        outputRow = outputRow + 1
    Next i
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateCurrentSOHReport"
    GoTo CleanExit

CleanExit:
    Set ws = Nothing
    Set qcCache = Nothing
    Set rmCache = Nothing
    Set rmCodeData = Nothing
    Set rmData = Nothing
    Exit Sub
End Sub

Private Sub AggregateAllSOHData( _
    ByVal qcCache As Object, _
    ByVal rmCache As Object, _
    ByVal rmCodeData As Object)
    
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim key As Variant
    Dim qcRecord As Object
    Dim rmRecord As Object
    Dim qcNumber As String
    Dim rmCode As String
    Dim rmCodeKey As Variant
    Dim stdName As String
    Dim masterUoM As String
    Dim category As String
    Dim quantity As Double
    Dim rmData As Object
    Dim activeQCNumbers As Object
    Dim grDateStr As String
    Dim nilStockFlag As String
    
    ' --- Object initialization ---
    Set activeQCNumbers = CreateObject("Scripting.Dictionary")
    
    ' --- Identify active QC numbers and build RM code structure ---
    For Each key In qcCache.keys
        Set qcRecord = qcCache(key)
        
        qcNumber = Trim$(Replace$(qcRecord(HEADER_QC_NUMBER), """", ""))
        grDateStr = Trim$(Replace$(qcRecord(HEADER_GR_DATE), """", ""))
        nilStockFlag = UCase$(Trim$(Replace$(qcRecord(HEADER_NIL_STOCK_FLAG), """", "")))
        
        If nilStockFlag = "Y" Then GoTo NextQCRecord
        
        activeQCNumbers(qcNumber) = True
        
        rmCode = Trim$(Replace$(qcRecord(HEADER_RM_CODE), """", ""))
        
        If Not rmCodeData.Exists(rmCode) Then
            stdName = Trim$(Replace$(qcRecord(HEADER_STANDARD_NAME), """", ""))
            masterUoM = Trim$(Replace$(qcRecord(HEADER_MASTER_UOM), """", ""))
            category = GetRMCategoryByCode(rmCode)
            
            Set rmData = CreateObject("Scripting.Dictionary")
            rmData("StandardName") = stdName
            rmData("MasterUoM") = masterUoM
            rmData("Category") = category
            rmData("TotalIns") = 0#
            rmData("TotalOuts") = 0#
            Set rmData("QCNumbers") = CreateObject("Scripting.Dictionary")
            rmCodeData.Add rmCode, rmData
        Else
            Set rmData = rmCodeData(rmCode)
        End If
        
        rmData("QCNumbers")(qcNumber) = True
        
NextQCRecord:
    Next key
    
    ' --- Aggregate RM Trans quantities ---
    For Each key In rmCache.keys
        Set rmRecord = rmCache(key)
        
        qcNumber = Trim$(Replace$(rmRecord(RMT_HEADER_QC_NUMBER), """", ""))
        
        If Not activeQCNumbers.Exists(qcNumber) Then GoTo NextRMRecord
        
        rmCode = ""
        For Each rmCodeKey In rmCodeData.keys
            Set rmData = rmCodeData(rmCodeKey)
            If rmData("QCNumbers").Exists(qcNumber) Then
                rmCode = CStr(rmCodeKey)
                Exit For
            End If
        Next rmCodeKey
        
        If rmCode = "" Then GoTo NextRMRecord
        
        Set rmData = rmCodeData(rmCode)
        
        On Error Resume Next
        quantity = CDbl(Trim$(Replace$(rmRecord(RMT_HEADER_QUANTITY), """", "")))
        If Err.Number <> 0 Then
            Err.Clear
            quantity = 0#
        End If
        On Error GoTo ErrHandler
        
        If quantity > 0# Then
            rmData("TotalIns") = rmData("TotalIns") + quantity
        Else
            rmData("TotalOuts") = rmData("TotalOuts") + quantity
        End If
        
NextRMRecord:
    Next key
    
    GoTo CleanExit

ErrHandler:
    GoTo CleanExit

CleanExit:
    Set qcRecord = Nothing
    Set rmRecord = Nothing
    Set rmData = Nothing
    Set activeQCNumbers = Nothing
    Exit Sub
End Sub

Private Function GetRMCategoryByCode(ByVal rmCode As String) As String
    On Error GoTo ErrHandler
    
    Dim wsRM As Worksheet
    Dim lastRow As Long
    Dim rmCodeCol As Long
    Dim categoryCol As Long
    Dim i As Long
    Dim searchCode As String
    Dim cellCode As String
    
    Set wsRM = ThisWorkbook.Worksheets("RM_Master")
    
    rmCodeCol = FindHeaderCol(wsRM, 1, RMM_HEADER_RM_CODE)
    categoryCol = FindHeaderCol(wsRM, 1, "Category")
    
    If rmCodeCol = 0 Or categoryCol = 0 Then
        GetRMCategoryByCode = ""
        GoTo CleanExit
    End If
    
    lastRow = wsRM.Cells(wsRM.Rows.count, rmCodeCol).End(xlUp).row
    searchCode = RM_Pad4(Trim$(rmCode))
    
    For i = 2 To lastRow
        cellCode = RM_Pad4(Trim$(CStr(wsRM.Cells(i, rmCodeCol).value)))
        If cellCode = searchCode Then
            GetRMCategoryByCode = Trim$(CStr(wsRM.Cells(i, categoryCol).value))
            GoTo CleanExit
        End If
    Next i
    
    GetRMCategoryByCode = ""
    GoTo CleanExit

ErrHandler:
    GetRMCategoryByCode = ""
    GoTo CleanExit

CleanExit:
    Set wsRM = Nothing
    Exit Function
End Function


