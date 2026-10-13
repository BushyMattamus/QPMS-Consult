Attribute VB_Name = "BusLogRequiresTesting"
Public Sub NewTesting()

    DisableAppUpdates
    
    Dim wsSource As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim reportRange As Range
    Dim qcTested As String
    Dim lastRetestDate As String
    Dim qcStatusFlag As String
    Dim i As Long, outRow As Long
    
    On Error GoTo ErrorHandler

    ' Fresh load the QC Number Data
    LoadQCNumberData
    
    ' Set worksheet references
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("QC_Number")
    Set wsReport = ThisWorkbook.Worksheets("New_Testing")
    On Error GoTo ErrorHandler
    
    ' Get last row of data in source
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No QC Number data found. Please refresh QC Number data first.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Clear existing content
    wsReport.Cells.Clear
    
    ' Write headers (expanded to include all required columns)
    wsReport.Range("B1:L1").value = Array( _
        HEADER_QC_NUMBER, _
        HEADER_RM_CODE, _
        HEADER_STANDARD_NAME, _
        HEADER_MANUFACTURER_LOT, _
        HEADER_MANUFACTURER_EXPIRY, _
        HEADER_NUMBER_CONTAINERS, _
        HEADER_TOTAL_QTY, _
        HEADER_QC_STATUS_FLAG, _
        HEADER_GR_DATE, _
        HEADER_SAMPLED_BY, _
        HEADER_SAMPLE_DATE)
    
    ' Copy filtered data (only "Sampled" status)
    outRow = 2
    For i = 2 To lastRow
    
        qcTested = Trim$(wsSource.Cells(i, 39).value) ' QC_Tested
        qcStatusFlag = LCase$(Trim$(wsSource.Cells(i, 34).value)) ' QC_Status_Flag
        lastRetestDate = Trim$(wsSource.Cells(i, 29).value) ' Last_Retest_Date
        
        
        If qcStatusFlag = "sampled" And (qcTested = "" Or qcTested = "No") And (lastRetestDate = "") Then
            wsReport.Cells(outRow, 2).value = wsSource.Cells(i, 1).value   ' QC_Number
            wsReport.Cells(outRow, 3).NumberFormat = "@"
            wsReport.Cells(outRow, 3).value = Right$("0000" & Trim$(wsSource.Cells(i, 2).text), 4)  ' RM_Code
            wsReport.Cells(outRow, 4).value = wsSource.Cells(i, 3).value   ' Standard_Name
            wsReport.Cells(outRow, 5).value = wsSource.Cells(i, 9).value   ' Mfg_Lot
            SetDateStringInCell wsReport.Cells(outRow, 6), Trim$(wsSource.Cells(i, 10).value) ' Mfg_Expiry
            wsReport.Cells(outRow, 7).value = wsSource.Cells(i, 11).value  ' Number_Containers
            wsReport.Cells(outRow, 8).value = wsSource.Cells(i, 12).value  ' Total_Qty
            wsReport.Cells(outRow, 9).value = wsSource.Cells(i, 34).value  ' QC_Status_Flag
            SetDateStringInCell wsReport.Cells(outRow, 10), Trim$(wsSource.Cells(i, 15).value) ' GR_Date
            wsReport.Cells(outRow, 11).value = wsSource.Cells(i, 19).value ' Sampled_By
            SetDateStringInCell wsReport.Cells(outRow, 12), Trim$(wsSource.Cells(i, 20).value) ' Sample_Date
            outRow = outRow + 1
        End If
    Next i
    
    ' Format the report range (11 columns: B to L)
    Set reportRange = wsReport.Range("B1").Resize(outRow - 1, 11)
    With reportRange
        .Columns(1).NumberFormat = "@"           ' QC Number as text
        .Columns(2).NumberFormat = "@"           ' RM Code as text
        .Columns(4).NumberFormat = "@"           ' Mfg_Lot as text
        .Columns(5).NumberFormat = "dd/mm/yyyy"  ' Mfg_Expiry as text (via SetDateStringInCell)
        .Columns(5).HorizontalAlignment = xlRight ' Right align date
        .Columns(9).NumberFormat = "dd/mm/yyyy"  ' GR_Date as text (via SetDateStringInCell)
        .Columns(9).HorizontalAlignment = xlRight ' Right align date
        .Columns(10).NumberFormat = "@"          ' Sampled By as text
        .Columns(11).NumberFormat = "dd/mm/yyyy" ' Sample_Date as text (via SetDateStringInCell)
        .Columns(11).HorizontalAlignment = xlRight ' Right align date
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Rows(1).Font.Bold = True
        .Locked = False
        .AutoFilter
        .EntireColumn.AutoFit
    End With

    wsReport.Range("B1:L1").Locked = False

    ' Check if any records were found
    If outRow = 2 Then
        MsgBox "No records found with status 'Sampled'.", vbInformation
        GoTo CleanExit
    End If

    ' Activate the report sheet
    wsReport.Activate
    wsReport.Range("B1").Select
    
    MsgBox "New Testing report created successfully. Found " & (outRow - 2) & " records requiring testing.", vbInformation
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error creating New Testing report: " & Err.Description, vbCritical
    
CleanExit:
End Sub

Public Sub UpcomingRetests()

    Dim wsSource As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim i As Long, outRow As Long
    Dim lastRetestDate As String
    Dim nilStockFlag As String
    Dim qcStatusFlag As String
    Dim testingFlag As String
    Dim extendedRetest As String
    Dim QCBalance As Double
    Dim today As Date, reviewPeriodEndDate As Date, sixMonthsEndDate As Date
    Dim reviewPeriodMonths As Double
    Dim reviewPeriodDays As Long

    On Error GoTo ErrorHandler

    LoadQCNumberData

    Set wsSource = ThisWorkbook.Worksheets("QC_Number")
    Set wsReport = ThisWorkbook.Worksheets("Upcoming_Retests")
    
    ' Get review period from user input (in months)
    reviewPeriodMonths = 0
    On Error Resume Next
    reviewPeriodMonths = CDbl(wsReport.Range("D3").value)
    On Error GoTo ErrorHandler
    
    If reviewPeriodMonths <= 0 Then
        MsgBox "Please enter a valid review period (in months) in cell D3.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Convert months to days (assuming 30 days per month)
    reviewPeriodDays = CLng(reviewPeriodMonths * 30)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No QC Number data found. Please refresh QC Number data first.", vbExclamation
        GoTo CleanExit
    End If

    ' Clear existing report content (but preserve user input area)
    wsReport.Range("F1:N" & wsReport.Rows.Count).Clear

    ' Write headers starting at F1
    wsReport.Range("F1:N1").value = Array( _
        HEADER_QC_NUMBER, _
        HEADER_RM_CODE, _
        HEADER_STANDARD_NAME, _
        "Upcoming Retest Date", _
        HEADER_MANUFACTURER_EXPIRY, _
        HEADER_QC_BALANCE, _
        HEADER_QC_STATUS_FLAG, _
        HEADER_NIL_STOCK_FLAG, _
        HEADER_EXTENDED_RETEST)

    today = Date
    reviewPeriodEndDate = DateAdd("d", reviewPeriodDays, today)
    sixMonthsEndDate = DateAdd("d", 183, today)
    outRow = 2

    For i = 2 To lastRow
        qcStatusFlag = LCase$(Trim$(wsSource.Cells(i, 34).value)) ' QC_Status_Flag
        testingFlag = UCase$(Trim$(wsSource.Cells(i, 22).value)) ' Testing_Flag
        nilStockFlag = UCase$(Trim$(wsSource.Cells(i, 31).value)) ' Nil_Stock_Flag
        lastRetestDate = Trim$(wsSource.Cells(i, 29).value) ' Last_Retest_Date
        extendedRetest = UCase$(Trim$(wsSource.Cells(i, 41).value)) ' Extended_Retest
        QCBalance = Val(wsSource.Cells(i, 21).value) ' QC_Balance (column 21)
        
        If (qcStatusFlag = "released" Or qcStatusFlag = "sampled" Or qcStatusFlag = "quarantine") And _
           (nilStockFlag = "" Or nilStockFlag = "N") And QCBalance > 0 And testingFlag = "Y" Then

            If lastRetestDate <> "" Then
                ' Parse the date for comparison (assuming DD/MM/YYYY format)
                Dim retestDateValue As Date
                On Error Resume Next
                retestDateValue = CDate(lastRetestDate)
                On Error GoTo ErrorHandler
                
                If (retestDateValue <= reviewPeriodEndDate And retestDateValue > 0) Or _
                    (extendedRetest = "Y" And retestDateValue <= sixMonthsEndDate) Then
                    wsReport.Cells(outRow, 6).value = wsSource.Cells(i, 1).value  ' QC_Number (col F)
                    wsReport.Cells(outRow, 7).NumberFormat = "@"
                    wsReport.Cells(outRow, 7).value = Right$("0000" & Trim$(wsSource.Cells(i, 2).text), 4)  ' RM_Code (col G)
                    wsReport.Cells(outRow, 8).value = wsSource.Cells(i, 3).value  ' Standard_Name (col H)
                    
                    ' Copy dates directly - already in DD/MM/YYYY format
                    SetDateStringInCell wsReport.Cells(outRow, 9), lastRetestDate ' Last_Retest_Date (col I)
                    SetDateStringInCell wsReport.Cells(outRow, 10), Trim$(wsSource.Cells(i, 10).value) ' Manufacturer_Expiry (col J)
                    
                    wsReport.Cells(outRow, 11).value = QCBalance ' QC_Balance (col K)
                    wsReport.Cells(outRow, 12).value = wsSource.Cells(i, 34).value ' QC_Status_Flag (col L)
                    wsReport.Cells(outRow, 13).value = wsSource.Cells(i, 31).value ' Nil_Stock_Flag (col M)
                    wsReport.Cells(outRow, 14).value = wsSource.Cells(i, 41).value ' Extended_Retest (col N)
                                      
                    ' Check if Extended Retest = "Y" and within 6 months - apply amber background
                    If extendedRetest = "Y" And retestDateValue <= sixMonthsEndDate Then
                        wsReport.Range(wsReport.Cells(outRow, 6), wsReport.Cells(outRow, 14)).Interior.Color = RGB(255, 191, 0) ' Amber/Orange color
                    End If
                    
                    ' Color retest dates that are equal to or older than today
                    If retestDateValue <= Date Then
                        wsReport.Cells(outRow, 9).Interior.Color = RGB(255, 124, 128) ' Red color for overdue retests
                    End If
                    
                    outRow = outRow + 1
                End If
            End If
        End If
    Next i

    ' Format the report range (starting at F1, 9 columns wide)
    Dim reportRange As Range
    Set reportRange = wsReport.Range("F1").Resize(outRow - 1, 9)
    With reportRange
        .Columns(1).NumberFormat = "@"           ' QC Number as text
        .Columns(2).NumberFormat = "@"           ' RM Code as text
        .Columns(4).NumberFormat = "dd/mm/yyyy"  ' Last Retest Date
        .Columns(4).HorizontalAlignment = xlRight ' Right align date
        .Columns(5).NumberFormat = "dd/mm/yyyy"  ' Manufacturer Expiry
        .Columns(5).HorizontalAlignment = xlRight ' Right align date
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Rows(1).Font.Bold = True
        .Locked = False
        .AutoFilter
        .EntireColumn.AutoFit
    End With

    wsReport.Range("F1:N1").Locked = False

    ' Sort by Extended Retest (priority), then by Last Retest Date
    If outRow > 2 Then
        ' First, convert text dates to actual date values for proper sorting
        Dim sortRange As Range
        Set sortRange = wsReport.Range("I2:I" & outRow - 1)
        
        ' Convert text dates to date values temporarily for sorting
        Dim cell As Range
        For Each cell In sortRange
            If cell.value <> "" Then
                On Error Resume Next
                Dim tempDate As Date
                tempDate = CDate(cell.value)
                If Err.Number = 0 Then
                    cell.value = tempDate
                End If
                On Error GoTo ErrorHandler
            End If
        Next cell
        
        wsReport.Sort.SortFields.Clear
        ' Primary sort: Extended Retest (descending - Y comes first)
        wsReport.Sort.SortFields.Add Key:=wsReport.Range("N2:N" & outRow - 1), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        ' Secondary sort: Last Retest Date (ascending)
        wsReport.Sort.SortFields.Add Key:=wsReport.Range("I2:I" & outRow - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With wsReport.Sort
            .SetRange wsReport.Range("F1:N" & outRow - 1)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' Convert dates back to DD/MM/YYYY format after sorting
        For Each cell In sortRange
            If IsDate(cell.value) Then
                cell.NumberFormat = "dd/mm/yyyy"
                cell.HorizontalAlignment = xlRight
            End If
        Next cell
    End If

    If outRow = 2 Then
        MsgBox "No upcoming retests found within the next " & reviewPeriodMonths & " months (" & reviewPeriodDays & " days).", vbInformation
        GoTo CleanExit
    End If

    wsReport.Activate
    wsReport.Range("F1").Select

    MsgBox "Upcoming Retests report created successfully. Found " & (outRow - 2) & " records within " & reviewPeriodMonths & " months.", vbInformation
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error creating Upcoming Retests: " & Err.Description, vbCritical

CleanExit:
End Sub


