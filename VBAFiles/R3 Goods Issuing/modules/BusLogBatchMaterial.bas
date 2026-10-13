Attribute VB_Name = "BusLogBatchMaterial"
' filepath: c:\Users\mattc\Documents\GitHub\CPISU-14-RM\R3 Goods Issuing\BusLogBatchMaterial.bas
Option Explicit

' ========================================================================
' BATCH MATERIALS USAGE REPORT
' ========================================================================
' Generates a report showing all raw materials issued to a specific
' finished goods batch using ADO SQL queries to join RM_Trans with
' QC Number and Raw Material Master data.
' ========================================================================

Public Sub GenerateBatchMaterialsReport()
    On Error GoTo ErrHandler

    Dim transactionCount As Long
    Dim materialCount As Long
    Dim conn As Object
    Dim rs As Object
    Dim rmTransPath As String
    Dim fgBatchMasterPath As String
    Dim qcCSVPath As String
    Dim rmMasterPath As String
    Dim tablesFolder As String
    Dim masterTablesFolder As String
    Dim batchNumber As String
    Dim batchInfo As Variant
    Dim outputData As Collection
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ' --- Get batch number from form ---
    batchNumber = Trim(GetFieldValByLabelC("Batch Number:"))
    
    If batchNumber = "" Then
        MsgBox "Please enter a Batch Number.", vbExclamation, "Missing Batch Number"
        GoTo CleanExit
    End If
    
    ' --- Build file paths ---
    tablesFolder = GetWorkbookTablesPath()
    masterTablesFolder = GetMasterDataPath(FOLDER_TABLES & "\")
    rmTransPath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    fgBatchMasterPath = GetWorkbookTableFilePath(FILE_FINISHED_GOODS_BATCH_MASTER)
    qcCSVPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    rmMasterPath = masterTablesFolder & FILE_RM_MASTER
    
    ' --- Validate file existence ---
    If Not FileExists(rmTransPath) Then
        MsgBox "RM_Trans.csv file not found at: " & rmTransPath, vbCritical, "File Missing"
        GoTo CleanExit
    End If
    If Not FileExists(fgBatchMasterPath) Then
        MsgBox "FinishedGoodsBatchNumberMasterFile.csv file not found at: " & fgBatchMasterPath, vbCritical, "File Missing"
        GoTo CleanExit
    End If
    If Not FileExists(qcCSVPath) Then
        MsgBox "QC_Number.csv file not found at: " & qcCSVPath, vbCritical, "File Missing"
        GoTo CleanExit
    End If
    If Not FileExists(rmMasterPath) Then
        MsgBox "Raw_Material_Master.csv file not found at: " & rmMasterPath, vbCritical, "File Missing"
        GoTo CleanExit
    End If
    
    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(tablesFolder)
    If conn Is Nothing Then
        MsgBox "Failed to open ADO connection.", vbCritical, "Connection Error"
        GoTo CleanExit
    End If
    
    ' --- Get batch info ---
    batchInfo = GetBatchInfoSQL(conn, batchNumber)
    
    If IsEmpty(batchInfo) Then
        MsgBox "Batch Number " & batchNumber & " not found in FinishedGoodsBatchNumberMaster.", vbInformation, "Batch Not Found"
        GoTo CleanExit
    End If
    
    ' --- Build materials data ---
    Set outputData = BuildBatchMaterialsDataSQL(conn, batchNumber)
    
    If outputData.Count = 0 Then
        MsgBox "No raw materials found for Batch Number " & batchNumber & ".", vbInformation, "No Materials Found"
        GoTo CleanExit
    End If
    
    ' --- Output report ---
    OutputBatchMaterialsReport outputData, batchInfo, ws
    
    transactionCount = 0
    materialCount = 0
    Dim tempRow As Variant
    Dim i As Long
    
    For i = 1 To outputData.Count
        tempRow = outputData(i)
        If tempRow(11) = "TOTAL" Then
            materialCount = materialCount + 1
        Else
            transactionCount = transactionCount + 1
        End If
    Next i
    
    MsgBox "Batch Materials report generated successfully." & vbCrLf & _
           "Batch: " & batchNumber & vbCrLf & _
           "Total materials: " & materialCount & vbCrLf & _
           "Total transactions: " & transactionCount, vbInformation, "Success"
    
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateBatchMaterialsReport"
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set outputData = Nothing
    Set ws = Nothing
    On Error GoTo 0
    Exit Sub
End Sub

Private Function GetBatchInfoSQL(ByVal conn As Object, ByVal batchNumber As String) As Variant
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim sql As String
    Dim info(1 To 3) As Variant
    Dim batchNumberField As String
    
    ' --- Detect actual first column name (handle BOM) ---
    batchNumberField = GetFirstColumnName(conn, FILE_FINISHED_GOODS_BATCH_MASTER)
    
    sql = "SELECT TOP 1 [" & FGBNM_HEADER_PMI_NUMBER & "], " & _
                 "[" & FGBNM_HEADER_PRODUCT_DESCRIPTION & "], " & _
                 "[" & FGBNM_HEADER_BATCH_NUMBER & "] " & _
          "FROM [" & FILE_FINISHED_GOODS_BATCH_MASTER & "] " & _
          "WHERE [" & FGBNM_HEADER_BATCH_NUMBER & "] = '" & EscapeSQL(batchNumber) & "' " & _
          "ORDER BY [" & FGBNM_HEADER_DATE_CREATED & "] DESC"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1
    
    If Not rs.EOF Then
        info(1) = Nz(rs.Fields(0).value, "")
        info(2) = Nz(rs.Fields(1).value, "")
        info(3) = Nz(rs.Fields(2).value, "")
        GetBatchInfoSQL = info
    Else
        GetBatchInfoSQL = Empty
    End If
    
    GoTo CleanExit

ErrHandler:
    GetBatchInfoSQL = Empty
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Function

Private Function BuildBatchMaterialsDataSQL(ByVal conn As Object, _
                                            ByVal batchNumber As String) As Collection
    On Error GoTo ErrHandler
    
    Dim results As Collection
    Dim rs As Object
    Dim sql As String
    Dim outputRow As Variant
    Dim qcNumberField_QC As String
    Dim qcNumberField_Trans As String
    Dim currentRMCode As String
    Dim previousRMCode As String
    Dim previousStdName As String
    Dim rmCodeTotal As Double
    
    Set results = New Collection
    
    ' --- Detect actual field names (handle BOM in first column) ---
    qcNumberField_QC = GetFirstColumnName(conn, FILE_QC_NUMBER)
    qcNumberField_Trans = RMT_HEADER_QC_NUMBER
    
    ' --- Build SQL with JOIN and aggregation ---
    ' Use subquery to get latest transaction details per QC_Number
    sql = "SELECT " & _
              "q.[" & HEADER_RM_CODE & "], " & _
              "q.[" & HEADER_STD_NAME & "], " & _
              "t.[" & qcNumberField_Trans & "], " & _
              "q.[" & HEADER_MANUFACTURER_LOT & "], " & _
              "q.[" & HEADER_LAST_RETEST_DATE & "], " & _
              "t.[" & RMT_HEADER_QUANTITY & "], " & _
              "q.[" & HEADER_MASTER_UOM & "], " & _
              "t.[" & RMT_HEADER_MEASURED_BY & "], " & _
              "t.[" & RMT_HEADER_USERNAME & "], " & _
              "t.[" & RMT_HEADER_TRANS_DT & "], " & _
              "t.[" & RMT_HEADER_ADJ_NOTE & "] " & _
          "FROM [" & FILE_RM_TRANS & "] AS t " & _
          "INNER JOIN [" & FILE_QC_NUMBER & "] AS q " & _
              "ON t.[" & qcNumberField_Trans & "] = q.[" & qcNumberField_QC & "] " & _
          "WHERE TRIM(t.[" & RMT_HEADER_FG_BATCH & "]) = '" & EscapeSQL(batchNumber) & "' " & _
          "ORDER BY q.[" & HEADER_RM_CODE & "], t.[" & RMT_HEADER_TRANS_DT & "]"
        
    Set rs = CreateObject("ADODB.Recordset")
    
    If conn Is Nothing Or conn.State = 0 Then
        MsgBox "Connection is not open.", vbCritical, "Connection Error"
        GoTo CleanExit
    End If
    
    On Error Resume Next
    rs.Open sql, conn, 1, 1
    If Err.Number <> 0 Then
        Debug.Print "Error " & Err.Number & ": " & Err.Description
        MsgBox "Failed to execute SQL query." & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbCritical, "SQL Execution Error"
        Err.Clear
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler
    
    previousRMCode = ""
    previousStdName = ""
    rmCodeTotal = 0
    
    Do While Not rs.EOF
        currentRMCode = StripBOM(Nz(rs.Fields(0).value, ""))
        
        If previousRMCode <> "" And currentRMCode <> previousRMCode Then
            ReDim outputRow(1 To 11)
            outputRow(1) = previousRMCode
            outputRow(2) = previousStdName
            outputRow(3) = ""
            outputRow(4) = ""
            outputRow(5) = ""
            outputRow(6) = rmCodeTotal
            outputRow(7) = ""
            outputRow(8) = ""
            outputRow(9) = ""
            outputRow(10) = ""
            outputRow(11) = "TOTAL"
            results.Add outputRow
            
            rmCodeTotal = 0
        End If
        
        ReDim outputRow(1 To 11)
        outputRow(1) = currentRMCode
        outputRow(2) = Nz(rs.Fields(1).value, "")
        outputRow(3) = Nz(rs.Fields(2).value, "")
        outputRow(4) = Nz(rs.Fields(3).value, "")
        outputRow(5) = Nz(rs.Fields(4).value, "")
        outputRow(6) = -CDbl(Nz(rs.Fields(5).value, 0))
        outputRow(7) = Nz(rs.Fields(6).value, "")
        outputRow(8) = Nz(rs.Fields(7).value, "")
        outputRow(9) = Nz(rs.Fields(8).value, "")
        outputRow(10) = FormatTransactionDate(Nz(rs.Fields(9).value, ""))
        outputRow(11) = Nz(rs.Fields(10).value, "")
        
        results.Add outputRow
        
        rmCodeTotal = rmCodeTotal + CDbl(Nz(rs.Fields(5).value, 0))
        previousRMCode = currentRMCode
        previousStdName = Nz(rs.Fields(1).value, "")
        
        rs.MoveNext
    Loop
    
    If previousRMCode <> "" Then
        ReDim outputRow(1 To 11)
        outputRow(1) = previousRMCode
        outputRow(2) = previousStdName
        outputRow(3) = ""
        outputRow(4) = ""
        outputRow(5) = ""
        outputRow(6) = rmCodeTotal
        outputRow(7) = ""
        outputRow(8) = ""
        outputRow(9) = ""
        outputRow(10) = ""
        outputRow(11) = "TOTAL"
        results.Add outputRow
    End If
    
    Set BuildBatchMaterialsDataSQL = results
    GoTo CleanExit

ErrHandler:
    Set BuildBatchMaterialsDataSQL = New Collection
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Function

Private Function GetFirstColumnName(ByVal conn As Object, ByVal fileName As String) As String
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim sql As String
    
    sql = "SELECT TOP 1 * FROM [" & fileName & "]"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1
    
    If rs.Fields.Count > 0 Then
        GetFirstColumnName = rs.Fields(0).Name
    Else
        GetFirstColumnName = ""
    End If
    
    GoTo CleanExit

ErrHandler:
    GetFirstColumnName = ""
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Function

Private Function EscapeSQL(ByVal text As String) As String
    EscapeSQL = Replace(text, "'", "''")
End Function

Private Function Nz(ByVal value As Variant, ByVal ValueIfNull As Variant) As Variant
    If IsNull(value) Or IsEmpty(value) Then
        Nz = ValueIfNull
    Else
        Nz = value
    End If
End Function

Private Sub OutputBatchMaterialsReport(ByVal data As Collection, ByVal batchInfo As Variant, ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim j As Long
    Dim outputRow As Variant
    Dim startRow As Long
    Dim lastRow As Long
    Dim endRow As Long
    Dim printRange As Range
    Dim dataRange As Range
    Dim reportHeading As String
    
    startRow = 6
    
    ' --- Set report header ---
    reportHeading = batchInfo(3) & ": " & batchInfo(1) & ", " & batchInfo(2)
    ws.Range("F4").value = reportHeading
    ws.Range("K4").value = "Reported by: " & Environ("USERNAME")

    ' --- Clear existing data ---
    ws.Range("F6:P100").ClearFormats
    
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    If lastRow >= startRow Then
        ws.Range(ws.Cells(startRow, 6), ws.Cells(lastRow, 16)).ClearContents
    End If
   
    ' --- Output data rows ---
    For i = 1 To data.Count
        outputRow = data(i)
        
        For j = 1 To 11
            If j = 1 Then
                With ws.Cells(startRow + i - 1, 5 + j)
                    .NumberFormat = "@"
                    .value = RM_Pad4(CStr(outputRow(j)))
                End With
            ElseIf j = 10 Then
                With ws.Cells(startRow + i - 1, 5 + j)
                    .NumberFormat = "@"
                    .value = outputRow(j)
                    .HorizontalAlignment = xlRight
                End With
            ElseIf j = 11 Then
                With ws.Cells(startRow + i - 1, 5 + j)
                    .value = outputRow(j)
                    .WrapText = True
                End With
            Else
                ws.Cells(startRow + i - 1, 5 + j).value = outputRow(j)
            End If
        Next j
    Next i
    
    ' --- Format output range ---
    endRow = startRow + data.Count - 1
    
    Set dataRange = ws.Range(ws.Cells(startRow, 6), ws.Cells(endRow, 16))
    
    With ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7))
        .WrapText = True
    End With
    
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' --- Apply double-line borders to total rows ---
    For i = 1 To data.Count
        outputRow = data(i)
        If outputRow(11) = "TOTAL" Then
            With ws.Range(ws.Cells(startRow + i - 1, 6), ws.Cells(startRow + i - 1, 16))
                .Font.Bold = True
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeBottom).Weight = xlThick
            End With
        End If
    Next i
    
    Set printRange = ws.Range("F4:P" & endRow)
    ws.PageSetup.PrintArea = printRange.Address

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "OutputBatchMaterialsReport"
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Set dataRange = Nothing
    Set printRange = Nothing
    On Error GoTo 0
End Sub

Private Function FormatTransactionDate(ByVal dateStr As String) As String
    On Error GoTo ErrHandler
    
    Dim datePart As String
    Dim dateParts() As String

    If Trim(dateStr) = "" Then
        FormatTransactionDate = ""
        GoTo CleanExit
    End If

    datePart = Trim(Split(dateStr, " ")(0))
    dateParts = Split(datePart, "-")
    
    If UBound(dateParts) = 2 Then
        FormatTransactionDate = Right$("0" & dateParts(2), 2) & "/" & _
                                Right$("0" & dateParts(1), 2) & "/" & _
                                dateParts(0)
    ElseIf IsDate(dateStr) Then
        FormatTransactionDate = Format(CDate(dateStr), "dd/mm/yyyy")
    Else
        FormatTransactionDate = dateStr
    End If

    GoTo CleanExit

ErrHandler:
    FormatTransactionDate = dateStr
    Resume CleanExit

CleanExit:
End Function

