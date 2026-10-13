Attribute VB_Name = "BusLogLiveReport"
Option Explicit

' Structure to hold row data for sorting
Private Type RowData
    qcNumber As String
    rmCode As String
    StandardName As String
    qcBalance As Variant
    QCStatus As String
    grDate As Variant
    grBy As String
    sampledBy As String
    sampleDate As Variant
    mfgLot As String
    mfgExpiry As Variant
End Type

Public Sub GenerateLiveQCReport()
    
    Dim wsSource As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim reportRange As Range
    Dim i As Long, outRow As Long
    
    ' Filter variables
    Dim rmDescription As String
    Dim rmCode As String
    Dim grDateStart As String
    Dim grDateEnd As String
    Dim grBy As String
    Dim sampledStartDate As String
    Dim sampledEndDate As String
    Dim sampledBy As String
    Dim QCStatus As String
    Dim nilStockFlag As String
    
    ' Date variables for comparisons
    Dim grStartDate As Date, grEndDate As Date
    Dim sampledStart As Date, sampledEnd As Date
    Dim hasGRStartDate As Boolean, hasGREndDate As Boolean
    Dim hasSampledStartDate As Boolean, hasSampledEndDate As Boolean
    
    ' Array for storing filtered data
    Dim filteredData() As RowData
    Dim dataCount As Long
    
    On Error GoTo ErrorHandler
    
    ' Get source worksheet
    Set wsSource = ThisWorkbook.Worksheets("QC_Number")
    Set wsReport = ActiveSheet
    
    ' Get last row of data in source
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then
        GoTo CleanExit
    End If
    
    ' Get filter values from form
    rmDescription = GetFieldValByLabelC("RM Description:")
    rmCode = GetFieldValByLabelC("RM Code:")
    grDateStart = GetFieldValByLabelC("GR Date Start:")
    grDateEnd = GetFieldValByLabelC("GR Date End:")
    grBy = GetFieldValByLabelC("GR By:")
    sampledStartDate = GetFieldValByLabelC("Sampled Start Date:")
    sampledEndDate = GetFieldValByLabelC("Sampled End Date:")
    sampledBy = GetFieldValByLabelC("Sampled By:")
    QCStatus = GetFieldValByLabelC("QC Status:")
    nilStockFlag = GetFieldValByLabelC("Nil Stock:")
    
    ' Parse dates if provided
    hasGRStartDate = False
    hasGREndDate = False
    hasSampledStartDate = False
    hasSampledEndDate = False
    
    If grDateStart <> "" And IsDate(grDateStart) Then
        grStartDate = CDate(grDateStart)
        hasGRStartDate = True
    End If
    
    If grDateEnd <> "" And IsDate(grDateEnd) Then
        grEndDate = CDate(grDateEnd)
        hasGREndDate = True
    End If
    
    If sampledStartDate <> "" And IsDate(sampledStartDate) Then
        sampledStart = CDate(sampledStartDate)
        hasSampledStartDate = True
    End If
    
    If sampledEndDate <> "" And IsDate(sampledEndDate) Then
        sampledEnd = CDate(sampledEndDate)
        hasSampledEndDate = True
    End If
    
    ' Clear existing report area (F3 onwards)
    wsReport.Range("F3:P1000").Clear
    
    ' Write headers starting at F3 (reordered columns with mfg fields)
    wsReport.Range("F3:P3").value = Array( _
        HEADER_RM_CODE, HEADER_STANDARD_NAME, HEADER_QC_NUMBER, _
        HEADER_MANUFACTURER_LOT, HEADER_MANUFACTURER_EXPIRY, HEADER_QC_BALANCE, _
        HEADER_QC_STATUS_FLAG, HEADER_GR_DATE, HEADER_GR_BY, _
        HEADER_SAMPLE_DATE, HEADER_SAMPLED_BY)

    ' Initialize filtered data array
    ReDim filteredData(1 To lastRow)
    dataCount = 0
    
    ' Filter data into array
    For i = 2 To lastRow
        If MatchesFilters(wsSource, i, rmDescription, rmCode, grBy, sampledBy, _
                         QCStatus, nilStockFlag, _
                         grStartDate, grEndDate, sampledStart, sampledEnd, _
                         hasGRStartDate, hasGREndDate, hasSampledStartDate, hasSampledEndDate) Then
            
            dataCount = dataCount + 1
            With filteredData(dataCount)
                .qcNumber = wsSource.Cells(i, 1).value
                .rmCode = wsSource.Cells(i, 2).value
                .StandardName = wsSource.Cells(i, 3).value
                .mfgLot = wsSource.Cells(i, 9).value
                .mfgExpiry = wsSource.Cells(i, 10).value
                .qcBalance = wsSource.Cells(i, 21).value
                .QCStatus = wsSource.Cells(i, 34).value
                .grDate = wsSource.Cells(i, 15).value
                .grBy = wsSource.Cells(i, 16).value
                .sampleDate = wsSource.Cells(i, 20).value
                .sampledBy = wsSource.Cells(i, 19).value
            End With
        End If
    Next i
    
    ' Sort the filtered data if we have any
    If dataCount > 0 Then
        ReDim Preserve filteredData(1 To dataCount)
        Call QuickSortRowData(filteredData, 1, dataCount)
        
        ' Output sorted data
        outRow = 4
        For i = 1 To dataCount
            With filteredData(i)
                wsReport.Cells(outRow, 6).value = .rmCode
                wsReport.Cells(outRow, 7).value = .StandardName
                wsReport.Cells(outRow, 8).value = .qcNumber
                wsReport.Cells(outRow, 9).value = .mfgLot
                wsReport.Cells(outRow, 10).value = .mfgExpiry
                wsReport.Cells(outRow, 11).value = .qcBalance
                wsReport.Cells(outRow, 12).value = .QCStatus
                wsReport.Cells(outRow, 13).value = .grDate
                wsReport.Cells(outRow, 14).value = .grBy
                wsReport.Cells(outRow, 15).value = .sampleDate
                wsReport.Cells(outRow, 16).value = .sampledBy
            End With
            outRow = outRow + 1
        Next i
    End If
    
    ' Format the report
    If outRow > 4 Then
        Set reportRange = wsReport.Range("F3").Resize(outRow - 3, 11)
        With reportRange
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Rows(1).Font.Bold = True
            .Locked = False
            .AutoFilter
            .EntireColumn.AutoFit
        End With
        
        ' Format RM_Code column (F) as text and apply 0000 format
        With wsReport.Range("F4:F" & (outRow - 1))
            .NumberFormat = "@"
            Dim j As Long
            For j = 4 To outRow - 1
                wsReport.Cells(j, 6).value = Format(wsReport.Cells(j, 6).value, "0000")
            Next j
        End With
    End If
    
    wsReport.Range("F3:P3").Locked = False
    
    GoTo CleanExit
    
ErrorHandler:
    ' Silent error handling for worksheet change events
    
CleanExit:
    Set wsSource = Nothing
    Set wsReport = Nothing
    Set sourceRange = Nothing
    Set reportRange = Nothing
End Sub

Private Sub QuickSortRowData(ByRef arr() As RowData, ByVal low As Long, ByVal high As Long)
    If low < high Then
        Dim pi As Long
        pi = PartitionRowData(arr, low, high)
        Call QuickSortRowData(arr, low, pi - 1)
        Call QuickSortRowData(arr, pi + 1, high)
    End If
End Sub

Private Function PartitionRowData(ByRef arr() As RowData, ByVal low As Long, ByVal high As Long) As Long
    Dim pivot As String
    Dim i As Long, j As Long
    
    ' Use RM Code as primary sort key, then Standard Name as secondary
    pivot = arr(high).rmCode & "|" & arr(high).StandardName
    i = low - 1
    
    For j = low To high - 1
        Dim compareValue As String
        compareValue = arr(j).rmCode & "|" & arr(j).StandardName
        
        If compareValue <= pivot Then
            i = i + 1
            Call SwapRowData(arr, i, j)
        End If
    Next j
    
    Call SwapRowData(arr, i + 1, high)
    PartitionRowData = i + 1
End Function

Private Sub SwapRowData(ByRef arr() As RowData, ByVal i As Long, ByVal j As Long)
    Dim temp As RowData
    temp = arr(i)
    arr(i) = arr(j)
    arr(j) = temp
End Sub

Private Function MatchesFilters(ByRef wsSource As Worksheet, ByVal row As Long, _
                               ByVal rmDescription As String, ByVal rmCode As String, _
                               ByVal grBy As String, ByVal sampledBy As String, _
                               ByVal QCStatus As String, ByVal nilStockFlag As String, _
                               ByVal grStartDate As Date, ByVal grEndDate As Date, _
                               ByVal sampledStart As Date, ByVal sampledEnd As Date, _
                               ByVal hasGRStartDate As Boolean, ByVal hasGREndDate As Boolean, _
                               ByVal hasSampledStartDate As Boolean, ByVal hasSampledEndDate As Boolean) As Boolean
    
    Dim cellValue As String
    Dim cellDate As Date
    
    ' Check RM Description filter (Standard Name - column 3)
    If rmDescription <> "" Then
        cellValue = Trim$(wsSource.Cells(row, 3).value)
        If InStr(1, cellValue, rmDescription, vbTextCompare) = 0 Then
            MatchesFilters = False
            Exit Function
        End If
    End If
    
    ' Check RM Code filter (column 2)
    If rmCode <> "" Then
        cellValue = Trim$(wsSource.Cells(row, 2).value)
        If StrComp(cellValue, Format(rmCode, "0000"), vbTextCompare) <> 0 Then
            MatchesFilters = False
            Exit Function
        End If
    End If
    
    ' Check GR Date Start filter (column 15)
    If hasGRStartDate Then
        cellValue = Trim$(wsSource.Cells(row, 15).value)
        If cellValue <> "" And IsDate(cellValue) Then
            cellDate = CDate(cellValue)
            If cellDate < grStartDate Then
                MatchesFilters = False
                Exit Function
            End If
        End If
    End If
    
    ' Check GR Date End filter (column 15)
    If hasGREndDate Then
        cellValue = Trim$(wsSource.Cells(row, 15).value)
        If cellValue <> "" And IsDate(cellValue) Then
            cellDate = CDate(cellValue)
            If cellDate > grEndDate Then
                MatchesFilters = False
                Exit Function
            End If
        End If
    End If
    
    ' Check GR By filter (column 16)
    If grBy <> "" Then
        cellValue = Trim$(wsSource.Cells(row, 16).value)
        If InStr(1, cellValue, grBy, vbTextCompare) = 0 Then
            MatchesFilters = False
            Exit Function
        End If
    End If
    
    ' Check Sampled Start Date filter (column 20)
    If hasSampledStartDate Then
        cellValue = Trim$(wsSource.Cells(row, 20).value)
        If cellValue <> "" And IsDate(cellValue) Then
            cellDate = CDate(cellValue)
            If cellDate < sampledStart Then
                MatchesFilters = False
                Exit Function
            End If
        End If
    End If
    
    ' Check Sampled End Date filter (column 20)
    If hasSampledEndDate Then
        cellValue = Trim$(wsSource.Cells(row, 20).value)
        If cellValue <> "" And IsDate(cellValue) Then
            cellDate = CDate(cellValue)
            If cellDate > sampledEnd Then
                MatchesFilters = False
                Exit Function
            End If
        End If
    End If
    
    ' Check Sampled By filter (column 19)
    If sampledBy <> "" Then
        cellValue = Trim$(wsSource.Cells(row, 19).value)
        If InStr(1, cellValue, sampledBy, vbTextCompare) = 0 Then
            MatchesFilters = False
            Exit Function
        End If
    End If
    
    ' Check QC Status filter (column 34)
    If QCStatus <> "" Then
        cellValue = Trim$(wsSource.Cells(row, 34).value)
        If StrComp(cellValue, QCStatus, vbTextCompare) <> 0 Then
            MatchesFilters = False
            Exit Function
        End If
    End If
    
    ' Check Nil Stock Flag filter (column 31)
    ' If user selects "N", show records with "N" or blank
    If nilStockFlag <> "" Then
        cellValue = Trim$(wsSource.Cells(row, 31).value)
        If UCase$(nilStockFlag) = "N" Then
            ' Accept "N" or blank/empty values
            If cellValue <> "" And StrComp(cellValue, "N", vbTextCompare) <> 0 Then
                MatchesFilters = False
                Exit Function
            End If
        Else
            ' For any other value (e.g., "Y"), require exact match
            If StrComp(cellValue, nilStockFlag, vbTextCompare) <> 0 Then
                MatchesFilters = False
                Exit Function
            End If
        End If
    End If
    
    ' If all filters pass, return True
    MatchesFilters = True
End Function
