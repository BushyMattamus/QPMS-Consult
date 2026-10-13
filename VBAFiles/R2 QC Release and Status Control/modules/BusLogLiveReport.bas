Attribute VB_Name = "BusLogLiveReport"
Option Explicit

' Structure to hold row data for sorting
Private Type RowData
    qcNumber As String
    rmCode As String
    StandardName As String
    VendorName As String
    MfgLot As String
    MfgExpiry As Variant
    NumberContainers As Variant
    TotalQty As Variant
    GRDate As Variant
    grBy As String
    sampledBy As String
    SampleDate As Variant
    QCBalance As Variant
    FirstReleaseDate As Variant
    FirstReleaseBy As String
    QCStatus As String
    ChangeReason As String
    TransDate As Variant
    TransUser As String
    NewRetest As String
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
    wsReport.Range("F3:Y1000").Clear
    
    ' Write headers starting at F3 (only the required fields)
    wsReport.Range("F3:Y3").value = Array( _
        HEADER_QC_NUMBER, _
        HEADER_RM_CODE, _
        HEADER_STANDARD_NAME, _
        HEADER_VENDOR_NAME, _
        HEADER_MANUFACTURER_LOT, _
        HEADER_MANUFACTURER_EXPIRY, _
        HEADER_NUMBER_CONTAINERS, _
        HEADER_TOTAL_QTY, _
        HEADER_GR_DATE, _
        HEADER_GR_BY, _
        HEADER_SAMPLED_BY, _
        HEADER_SAMPLE_DATE, _
        HEADER_QC_BALANCE, _
        HEADER_FIRST_RELEASE_DATE, _
        HEADER_FIRST_RELEASE_BY, _
        HEADER_QC_STATUS_FLAG, _
        HEADER_CHANGE_REASON, _
        HEADER_TRANS_DATE, _
        HEADER_TRANS_USER, _
        HEADER_NEW_RETEST)

    ' Initialize filtered data array
    ReDim filteredData(1 To lastRow)
    dataCount = 0
    
    ' Filter data into array
    For i = 2 To lastRow
        If MatchesFilters(wsSource, i, rmDescription, rmCode, grBy, sampledBy, _
                         grStartDate, grEndDate, sampledStart, sampledEnd, _
                         hasGRStartDate, hasGREndDate, hasSampledStartDate, hasSampledEndDate, _
                         QCStatus, nilStockFlag) Then
            
            dataCount = dataCount + 1
            With filteredData(dataCount)
                .qcNumber = wsSource.Cells(i, 1).value              ' QC_Number
                .rmCode = wsSource.Cells(i, 2).value                ' RM_Code
                .StandardName = wsSource.Cells(i, 3).value          ' Standard_Name
                .VendorName = wsSource.Cells(i, 7).value            ' Vendor_Name (col 7)
                .MfgLot = wsSource.Cells(i, 9).value                ' Mfg_Lot (col 9)
                .MfgExpiry = wsSource.Cells(i, 10).value            ' Mfg_Expiry (col 10)
                .NumberContainers = wsSource.Cells(i, 11).value     ' Number_Containers (col 11)
                .TotalQty = wsSource.Cells(i, 12).value             ' Total_Qty (col 12)
                .GRDate = wsSource.Cells(i, 15).value               ' GR_Date (col 15)
                .grBy = wsSource.Cells(i, 16).value                 ' GR_By (col 16)
                .sampledBy = wsSource.Cells(i, 19).value            ' Sampled_By (col 19)
                .SampleDate = wsSource.Cells(i, 20).value           ' Sample_Date (col 20)
                .QCBalance = wsSource.Cells(i, 21).value            ' QC_Balance (col 21)
                .FirstReleaseDate = wsSource.Cells(i, 23).value     ' First_Release_Date (col 23)
                .FirstReleaseBy = wsSource.Cells(i, 24).value       ' First_Release_by (col 24)
                .QCStatus = wsSource.Cells(i, 34).value             ' QC_Status_Flag (col 34)
                .ChangeReason = wsSource.Cells(i, 35).value         ' Change_Reason (col 35)
                .TransDate = wsSource.Cells(i, 36).value            ' Trans_Date (col 36)
                .TransUser = wsSource.Cells(i, 37).value            ' Trans_User (col 37)
                .NewRetest = wsSource.Cells(i, 38).value            ' New_Retest (col 38)
            End With
        End If
    Next i
    
    ' Sort the filtered data if we have any
    If dataCount > 0 Then
        ReDim Preserve filteredData(1 To dataCount)
        Call QuickSortRowData(filteredData, 1, dataCount)
        
        ' Output sorted data (20 columns starting at F)
        outRow = 4
        For i = 1 To dataCount
            With filteredData(i)
                wsReport.Cells(outRow, 6).value = .qcNumber                    ' F
                wsReport.Cells(outRow, 7).NumberFormat = "@"
                wsReport.Cells(outRow, 7).value = Right$("0000" & Trim$(.rmCode), 4)  ' G
                wsReport.Cells(outRow, 8).value = .StandardName                ' H
                wsReport.Cells(outRow, 9).value = .VendorName                  ' I
                wsReport.Cells(outRow, 10).value = .MfgLot                     ' J
                SetDateStringInCell wsReport.Cells(outRow, 11), Trim$(.MfgExpiry)  ' K
                wsReport.Cells(outRow, 12).value = .NumberContainers           ' L
                wsReport.Cells(outRow, 13).value = .TotalQty                   ' M
                SetDateStringInCell wsReport.Cells(outRow, 14), Trim$(.GRDate)     ' N
                wsReport.Cells(outRow, 15).value = .grBy                       ' O
                wsReport.Cells(outRow, 16).value = .sampledBy                  ' P
                SetDateStringInCell wsReport.Cells(outRow, 17), Trim$(.SampleDate) ' Q
                wsReport.Cells(outRow, 18).value = .QCBalance                  ' R
                SetDateStringInCell wsReport.Cells(outRow, 19), Trim$(.FirstReleaseDate) ' S
                wsReport.Cells(outRow, 20).value = .FirstReleaseBy             ' T
                wsReport.Cells(outRow, 21).value = .QCStatus                   ' U
                wsReport.Cells(outRow, 22).value = .ChangeReason               ' V
                SetDateStringInCell wsReport.Cells(outRow, 23), Trim$(.TransDate)  ' W
                wsReport.Cells(outRow, 24).value = .TransUser                  ' X
                ' Display "New" if New_Retest is empty, otherwise show the actual value
                If Trim$(.NewRetest) = "" Then
                    wsReport.Cells(outRow, 25).value = "New"                   ' Y
                Else
                    wsReport.Cells(outRow, 25).value = .NewRetest              ' Y
                End If
            End With
            outRow = outRow + 1
        Next i
    End If
    
    ' Format the report (20 columns: F to Y)
    If outRow > 4 Then
        Set reportRange = wsReport.Range("F3").Resize(outRow - 3, 20)
        With reportRange
            .Columns(1).NumberFormat = "@"           ' QC Number as text
            .Columns(2).NumberFormat = "@"           ' RM Code as text
            .Columns(5).NumberFormat = "@"           ' Mfg_Lot as text
            .Columns(6).NumberFormat = "@"           ' Mfg_Expiry as text
            .Columns(6).HorizontalAlignment = xlRight ' Right align date
            .Columns(9).NumberFormat = "@"           ' GR_Date as text
            .Columns(9).HorizontalAlignment = xlRight ' Right align date
            .Columns(12).NumberFormat = "@"          ' Sample_Date as text
            .Columns(12).HorizontalAlignment = xlRight ' Right align date
            .Columns(14).NumberFormat = "@"          ' First_Release_Date as text
            .Columns(14).HorizontalAlignment = xlRight ' Right align date
            .Columns(18).NumberFormat = "@"          ' Trans_Date as text
            .Columns(18).HorizontalAlignment = xlRight ' Right align date
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Rows(1).Font.Bold = True
            .Locked = False
            .AutoFilter
            .EntireColumn.AutoFit
        End With
    End If

    wsReport.Range("F3:Y3").Locked = False
    
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
    
    ' Use QC Number as primary sort key, then Standard Name as secondary
    pivot = arr(high).qcNumber & "|" & arr(high).StandardName
    i = low - 1
    
    For j = low To high - 1
        Dim compareValue As String
        compareValue = arr(j).qcNumber & "|" & arr(j).StandardName
        
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
                               ByVal grStartDate As Date, ByVal grEndDate As Date, _
                               ByVal sampledStart As Date, ByVal sampledEnd As Date, _
                               ByVal hasGRStartDate As Boolean, ByVal hasGREndDate As Boolean, _
                               ByVal hasSampledStartDate As Boolean, ByVal hasSampledEndDate As Boolean, _
                               ByVal QCStatus As String, ByVal nilStockFlag As String) As Boolean
    
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
        If InStr(1, cellValue, QCStatus, vbTextCompare) = 0 Then
            MatchesFilters = False
            Exit Function
        End If
    End If
    
    ' Check Nil Stock Flag filter (column 31)
    ' If "N" is selected, also return blank values
    If nilStockFlag <> "" Then
        cellValue = UCase$(Trim$(wsSource.Cells(row, 31).value))
        If UCase$(nilStockFlag) = "N" Then
            ' For "N", accept both "N" and blank values
            If cellValue <> "N" And cellValue <> "" Then
                MatchesFilters = False
                Exit Function
            End If
        Else
            ' For "Y" or any other value, must match exactly
            If cellValue <> UCase$(nilStockFlag) Then
                MatchesFilters = False
                Exit Function
            End If
        End If
    End If
    
    ' If all filters pass, return True
    MatchesFilters = True
End Function
