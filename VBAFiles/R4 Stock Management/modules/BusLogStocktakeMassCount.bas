Attribute VB_Name = "BusLogStocktakeMassCount"
Option Explicit

' ========================================================================
' STOCKTAKE MASS COUNT SHEET
' ========================================================================
' Handles bulk stocktake entry based on RM Category filter
' Generates list of uncounted QC Numbers and allows batch data entry
' ========================================================================

' --- Constants for column positions ---
Private Const COL_HEADER_RMCODE As Long = 8 ' Column H
Private Const COL_HEADER_STDNAME As Long = 9 ' Column I
Private Const COL_HEADER_QCNUM As Long = 10 ' Column J
Private Const COL_HEADER_LASTRETESTDATE As Long = 11 ' Column K
Private Const COL_HEADER_BOOKQTY As Long = 12 ' Column L
Private Const COL_HEADER_UOM As Long = 13 ' Column M
Private Const COL_HEADER_COUNTQTY As Long = 14 ' Column N
Private Const COL_HEADER_RECORDEDRETESTDATE As Long = 15 ' Column O
Private Const COL_HEADER_COUNTBY As Long = 16 ' Column P
Private Const COL_HEADER_DATECOUNTED As Long = 17 ' Column Q

Private Const ROW_HEADER As Long = 2 ' Headers on row 1
Private Const ROW_DATA_START As Long = 3 ' Data starts row 2

' ========================================================================
' Generate count list based on selected RM Category
' Called when user changes the RM Category dropdown in Column D
' ========================================================================
Public Sub GenerateStocktakeCountList()
    ' --- Variable declarations ---
    Dim rmCategory As String
    Dim productionFilter As String
    Dim countSheet As Worksheet
    Dim rmMasterSheet As Worksheet
    Dim qcSheet As Worksheet
    Dim rmCodesInCategory As Object
    Dim balanceLookup As Object
    Dim countedQCs As Object
    Dim lastCountDateLookup As Object
    Dim lastRow As Long
    Dim r As Long
    Dim writeRow As Long
    Dim qcNumber As String
    Dim rmCode As String
    Dim stdName As String
    Dim masterUoM As String
    Dim bookQty As Double
    Dim colRMCode As Long
    Dim colCategory As Long
    Dim colQC As Long
    Dim colRM As Long
    Dim colStd As Long
    Dim colUoM As Long
    Dim colNilStockFlagDate As Long
    Dim colFirstReleaseDate As Long
    Dim colFirstRetestDate As Long
    Dim colLastRetestDate As Long
    Dim qcCat As String
    Dim itemsAdded As Long
    Dim sortRange As Range
    Dim firstReleaseDate As String
    Dim firstRetestDate As String
    Dim lastRetestDate As String
    Dim nilStockFlagDate As String
    Dim lastCountDate As String
    Dim isInProduction As Boolean
    Dim useRMFilter As Boolean
    Dim useLocationFilter As Boolean

    On Error GoTo ErrHandler

    ' --- Refresh QC Number data ---
    BusLogDataLoads.LoadQCNumberData

    ClearStocktakeMassCountSheet

    Set countSheet = ActiveSheet
    Set rmMasterSheet = ThisWorkbook.Worksheets("RM_Master")
    Set qcSheet = ThisWorkbook.Worksheets("QC_Number")

    ' --- Get selected RM Category and Location (optional) ---
    rmCategory = Trim$(GetFieldValByLabelC("RM Category:"))
    productionFilter = Trim$(GetFieldValByLabelC("Location:"))

    useRMFilter = (rmCategory <> "")
    useLocationFilter = (productionFilter <> "")

    ' --- STEP 1: Get all RM_Codes that belong to selected category (if provided) ---
    Set rmCodesInCategory = CreateObject("Scripting.Dictionary")

    If useRMFilter Then
        colRMCode = FindHeaderCol(rmMasterSheet, 1, RMM_HEADER_RM_CODE)
        colCategory = FindHeaderCol(rmMasterSheet, 1, RMM_HEADER_CATEGORY)

        If colRMCode = 0 Or colCategory = 0 Then
            MsgBox "Required columns not found in RM_Master worksheet.", vbExclamation
            GoTo CleanExit
        End If

        lastRow = rmMasterSheet.Cells(rmMasterSheet.Rows.count, colRMCode).End(xlUp).row

        For r = 2 To lastRow
            qcCat = Trim$(CStr(rmMasterSheet.Cells(r, colCategory).value))
            
            If UCase$(qcCat) = UCase$(rmCategory) Then
                rmCode = Trim$(CStr(rmMasterSheet.Cells(r, colRMCode).value))
                If rmCode <> "" Then
                    rmCodesInCategory(rmCode) = True
                End If
            End If
        Next r

        If rmCodesInCategory.count = 0 Then
            MsgBox "No RM Codes found for category: " & rmCategory, vbExclamation
            GoTo CleanExit
        End If
    End If

    ' --- STEP 2: Load current balances and counted QCs ---
    Set balanceLookup = CreateObject("Scripting.Dictionary")
    Call LoadQCBalancesFromRMTrans(balanceLookup)

    Set countedQCs = GetStocktakeMonthStatus()
    Set lastCountDateLookup = GetLastCountDateLookup()

    ' --- STEP 3: Get QC_Number sheet columns ---
    colQC = FindHeaderCol(qcSheet, 1, HEADER_QC_NUMBER)
    colRM = FindHeaderCol(qcSheet, 1, HEADER_RM_CODE)
    colStd = FindHeaderCol(qcSheet, 1, HEADER_STANDARD_NAME)
    colUoM = FindHeaderCol(qcSheet, 1, HEADER_MASTER_UOM)
    colNilStockFlagDate = FindHeaderCol(qcSheet, 1, HEADER_NIL_STOCK_FLAG_DATE)
    colFirstReleaseDate = FindHeaderCol(qcSheet, 1, HEADER_FIRST_RELEASE_DATE)
    colFirstRetestDate = FindHeaderCol(qcSheet, 1, HEADER_FIRST_RETEST_DATE)
    colLastRetestDate = FindHeaderCol(qcSheet, 1, HEADER_LAST_RETEST_DATE)

    If colQC = 0 Then
        MsgBox "Required columns not found in QC_Number worksheet.", vbExclamation
        GoTo CleanExit
    End If

    ' --- STEP 4: Clear existing data rows (keep headers on row 4) ---
    lastRow = countSheet.Cells(countSheet.Rows.count, COL_HEADER_RMCODE).End(xlUp).row
    If lastRow >= ROW_DATA_START Then
        countSheet.Range(countSheet.Cells(ROW_DATA_START, COL_HEADER_RMCODE), _
                         countSheet.Cells(lastRow, COL_HEADER_DATECOUNTED)).ClearContents
    End If
    
    ' --- STEP 5: Loop through QC_Number sheet and apply filters ---
    writeRow = ROW_DATA_START
    lastRow = qcSheet.Cells(qcSheet.Rows.count, colQC).End(xlUp).row
    itemsAdded = 0
    
    For r = 2 To lastRow
        qcNumber = Trim$(CStr(qcSheet.Cells(r, colQC).value))
        rmCode = Trim$(CStr(qcSheet.Cells(r, colRM).value))
        
        If useRMFilter Then
            If Not rmCodesInCategory.Exists(rmCode) Then
                GoTo SkipRow
            End If
        End If
        
        If Not countedQCs.Exists(qcNumber) Then
            If balanceLookup.Exists(qcNumber) Then
                bookQty = balanceLookup(qcNumber)
            Else
                bookQty = 0
            End If
            
            ' --- Check nil stock exclusion logic ---
            If colNilStockFlagDate > 0 Then
                nilStockFlagDate = Trim$(CStr(qcSheet.Cells(r, colNilStockFlagDate).value))
                
                If nilStockFlagDate <> "" Then
                    If lastCountDateLookup.Exists(qcNumber) Then
                        lastCountDate = CStr(lastCountDateLookup(qcNumber))
                        If DateToANSI(lastCountDate) >= DateToANSI(nilStockFlagDate) Then
                            GoTo SkipRow
                        End If
                    End If
                End If
            End If
            
            isInProduction = False
            
            If colFirstReleaseDate > 0 Then
                firstReleaseDate = Trim$(CStr(qcSheet.Cells(r, colFirstReleaseDate).value))
                If firstReleaseDate <> "" Then
                    isInProduction = True
                End If
            End If
            
            If Not isInProduction And colFirstRetestDate > 0 Then
                firstRetestDate = Trim$(CStr(qcSheet.Cells(r, colFirstRetestDate).value))
                If firstRetestDate <> "" Then
                    isInProduction = True
                End If
            End If
            
            If useLocationFilter Then
                If UCase$(productionFilter) = "PRODUCTION" And Not isInProduction Then
                    GoTo SkipRow
                End If
                
                If UCase$(productionFilter) = "INBOUND" And isInProduction Then
                    GoTo SkipRow
                End If
            End If
            
            stdName = ""
            masterUoM = ""
            
            If colStd > 0 Then stdName = Trim$(CStr(qcSheet.Cells(r, colStd).value))
            If colUoM > 0 Then masterUoM = Trim$(CStr(qcSheet.Cells(r, colUoM).value))
            
            With countSheet
                .Cells(writeRow, COL_HEADER_RMCODE).NumberFormat = "@"
                .Cells(writeRow, COL_HEADER_RMCODE).value = RM_Pad4(rmCode)
                
                .Cells(writeRow, COL_HEADER_STDNAME).value = stdName
                .Cells(writeRow, COL_HEADER_QCNUM).value = qcNumber
                
                If colLastRetestDate > 0 Then
                    lastRetestDate = Trim$(CStr(qcSheet.Cells(r, colLastRetestDate).value))
                    If lastRetestDate <> "" Then
                        .Cells(writeRow, COL_HEADER_LASTRETESTDATE).NumberFormat = "dd/mm/yyyy"
                        If IsDate(lastRetestDate) Then
                            .Cells(writeRow, COL_HEADER_LASTRETESTDATE).value = CDate(lastRetestDate)
                        Else
                            .Cells(writeRow, COL_HEADER_LASTRETESTDATE).value = lastRetestDate
                        End If
                    End If
                End If
                
                .Cells(writeRow, COL_HEADER_BOOKQTY).value = Format$(bookQty, "0.000")
                .Cells(writeRow, COL_HEADER_UOM).value = masterUoM
                
                .Cells(writeRow, COL_HEADER_RMCODE).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_STDNAME).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_QCNUM).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_LASTRETESTDATE).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_BOOKQTY).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_UOM).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_COUNTQTY).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_RECORDEDRETESTDATE).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_COUNTBY).Borders.LineStyle = xlContinuous
                .Cells(writeRow, COL_HEADER_DATECOUNTED).Borders.LineStyle = xlContinuous
            End With
            
            itemsAdded = itemsAdded + 1
            writeRow = writeRow + 1
        End If
        
SkipRow:
    Next r
    
    ' --- STEP 7: Sort the data by RM Code, then by QC Number ---
    If itemsAdded > 0 Then
        Set sortRange = countSheet.Range(countSheet.Cells(ROW_DATA_START, COL_HEADER_RMCODE), _
                                         countSheet.Cells(writeRow - 1, COL_HEADER_DATECOUNTED))
        
        With countSheet.Sort
            .SortFields.Clear
            .SortFields.Add key:=countSheet.Cells(ROW_DATA_START, COL_HEADER_RMCODE), _
                           SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add key:=countSheet.Cells(ROW_DATA_START, COL_HEADER_QCNUM), _
                           SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange sortRange
            .header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' --- Autofit row heights ---
        countSheet.Rows(ROW_DATA_START & ":" & writeRow - 1).EntireRow.AutoFit
        
        ' --- Unlock the cells in the created range ---
        sortRange.Locked = False
        
    End If
    
    MsgBox "Count list generated successfully." & vbCrLf & _
           "Items to count: " & itemsAdded, _
           vbInformation, "Count List Ready"

    countSheet.Cells(ROW_DATA_START, COL_HEADER_COUNTQTY).Select
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateStocktakeCountList"
    
CleanExit:
    Set sortRange = Nothing
    Set countedQCs = Nothing
    Set balanceLookup = Nothing
    Set lastCountDateLookup = Nothing
    Set rmCodesInCategory = Nothing
    Set qcSheet = Nothing
    Set rmMasterSheet = Nothing
    Set countSheet = Nothing
    Exit Sub
End Sub

' ========================================================================
' Submit all counts from the mass entry sheet (PAYG - Post As You Go)
' ========================================================================
Public Sub SubmitStocktakeMassCountSheet()
    Dim countSheet As Worksheet
    Dim dataRow As Long
    Dim lastRow As Long
    Dim successCount As Long
    Dim failureCount As Long
    Dim noCount As Long
    Dim qcNumber As String
    Dim countQty As String
    Dim countBy As String
    Dim dateStr As String
    Dim countedQtyVal As Double
    Dim countedUoM As String
    Dim recordedRetestDate As String
    Dim ok As Boolean
    
    On Error GoTo ErrHandler
    
    Set countSheet = ActiveSheet
    
    lastRow = countSheet.Cells(countSheet.Rows.count, COL_HEADER_QCNUM).End(xlUp).row
    
    If lastRow < ROW_DATA_START Then
        MsgBox "No count data to submit. Please generate a count list first.", vbExclamation
        GoTo CleanExit
    End If
    
    If MsgBox("Submit all counts from this sheet?" & vbCrLf & vbCrLf & _
             "This cannot be undone.", vbYesNo + vbExclamation, "Confirm Submission") = vbNo Then
        GoTo CleanExit
    End If
        
    successCount = 0
    failureCount = 0
    noCount = 0
    
    For dataRow = ROW_DATA_START To lastRow
        qcNumber = Trim$(CStr(countSheet.Cells(dataRow, COL_HEADER_QCNUM).value))
        countQty = Trim$(CStr(countSheet.Cells(dataRow, COL_HEADER_COUNTQTY).value))
        countBy = Trim$(CStr(countSheet.Cells(dataRow, COL_HEADER_COUNTBY).value))
        dateStr = Trim$(CStr(countSheet.Cells(dataRow, COL_HEADER_DATECOUNTED).value))
        countedUoM = Trim$(CStr(countSheet.Cells(dataRow, COL_HEADER_UOM).value))
        recordedRetestDate = Trim$(CStr(countSheet.Cells(dataRow, COL_HEADER_RECORDEDRETESTDATE).value))
        
        If Not ValidateNoCommas(countBy, "Counted By") Then GoTo CleanExit
        If Not ValidateNoCommas(countQty, "Count Quantity") Then GoTo CleanExit
        
        If qcNumber = "" Then
            GoTo NextRow
        End If
        
        If countQty = "" Then
            noCount = noCount + 1
            GoTo NextRow
        End If
        
        If Not IsNumeric(countQty) Then
            failureCount = failureCount + 1
            GoTo NextRow
        End If
        
        If countBy = "" Then countBy = Environ$("Username")
        
        countedQtyVal = CDbl(countQty)
        
        ok = StocktakeQCNumEntry(qcNumber, countedQtyVal, countedUoM, countBy, "", recordedRetestDate)

        If ok Then
            successCount = successCount + 1
            countSheet.Cells(dataRow, COL_HEADER_COUNTQTY).Interior.Color = RGB(192, 255, 192)
        Else
            failureCount = failureCount + 1
            countSheet.Cells(dataRow, COL_HEADER_COUNTQTY).Interior.Color = RGB(255, 192, 192)
        End If
        
NextRow:
    Next dataRow
    
    Dim summaryMsg As String
    summaryMsg = "Stocktake submission complete." & vbCrLf & vbCrLf & _
                 "Successful: " & successCount & vbCrLf & _
                 "Failed: " & failureCount & vbCrLf & _
                 "No count: " & noCount
    
    If failureCount > 0 Then
        summaryMsg = summaryMsg & vbCrLf & vbCrLf & "See red-highlighted rows for failures."
    End If
    
    MsgBox summaryMsg, IIf(failureCount > 0, vbExclamation, vbInformation), "Submission Summary"
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & " at row " & dataRow & ": " & Err.Description, vbCritical, "SubmitStocktakeMassCountSheet"
    
CleanExit:
    Set countSheet = Nothing
    Exit Sub
End Sub

' ========================================================================
' Clear all count data from the sheet (keep headers)
' ========================================================================

Public Sub ClearStocktakeMassCountSheet()
    Dim countSheet As Worksheet
    Dim clearRange As Range
    Dim lastRow As Long
    
    On Error GoTo ErrHandler
    
    Set countSheet = ActiveSheet
    lastRow = countSheet.Cells(countSheet.Rows.count, COL_HEADER_RMCODE).End(xlUp).row
    
    If lastRow >= ROW_DATA_START Then
        Set clearRange = countSheet.Range(countSheet.Cells(ROW_DATA_START, COL_HEADER_RMCODE), _
                                          countSheet.Cells(lastRow, COL_HEADER_DATECOUNTED))
        clearRange.ClearContents
        clearRange.Borders.LineStyle = xlNone
        clearRange.Interior.ColorIndex = xlNone
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ClearStocktakeMassCountSheet"

CleanExit:
    Set clearRange = Nothing
    Set countSheet = Nothing
    Exit Sub
End Sub

' ========================================================================
' Generate summary of current month stocktake count status
' Populates Inbound and Production count summaries based on all eligible QC numbers
' ========================================================================
Public Sub GenerateStocktakeSummary()
    ' --- Variable declarations ---
    Dim qcSheet As Worksheet
    Dim countSheet As Worksheet
    Dim balanceLookup As Object
    Dim countedQCs As Object
    Dim inboundCounted As Long, inboundToCount As Long
    Dim productionCounted As Long, productionToCount As Long
    Dim lastRow As Long
    Dim r As Long
    Dim qcNumber As String
    Dim bookQty As Double
    Dim colQC As Long, colNilStock As Long
    Dim colFirstReleaseDate As Long, colFirstRetestDate As Long
    Dim nilStockFlag As String
    Dim firstReleaseDate As String, firstRetestDate As String
    Dim isInProduction As Boolean

    On Error GoTo ErrHandler

    Set qcSheet = ThisWorkbook.Worksheets("QC_Number")
    Set countSheet = ActiveSheet
    Set balanceLookup = CreateObject("Scripting.Dictionary")

    ' --- Refresh and load data ---
    BusLogDataLoads.LoadQCNumberData
    Call LoadQCBalancesFromRMTrans(balanceLookup)
    Set countedQCs = GetStocktakeMonthStatus()

    ' --- Get QC_Number sheet columns ---
    colQC = FindHeaderCol(qcSheet, 1, HEADER_QC_NUMBER)
    colNilStock = FindHeaderCol(qcSheet, 1, HEADER_NIL_STOCK_FLAG)
    colFirstReleaseDate = FindHeaderCol(qcSheet, 1, HEADER_FIRST_RELEASE_DATE)
    colFirstRetestDate = FindHeaderCol(qcSheet, 1, HEADER_FIRST_RETEST_DATE)

    If colQC = 0 Then
        MsgBox "Required column not found in QC_Number worksheet.", vbExclamation
        GoTo CleanExit
    End If

    ' --- Initialize counters ---
    inboundCounted = 0
    inboundToCount = 0
    productionCounted = 0
    productionToCount = 0

    ' --- Loop through QC_Number sheet and count all eligible QC numbers ---
    lastRow = qcSheet.Cells(qcSheet.Rows.count, colQC).End(xlUp).row

    For r = 2 To lastRow
        qcNumber = Trim$(CStr(qcSheet.Cells(r, colQC).value))

        ' --- Get balance and check exclusion criteria ---
        If balanceLookup.Exists(qcNumber) Then
            bookQty = balanceLookup(qcNumber)
        Else
            bookQty = 0
        End If

        ' --- Exclude if balance is 0 or nil stock flag is Y ---
        If bookQty = 0 Then
            GoTo SkipRow
        End If

        If colNilStock > 0 Then
            nilStockFlag = Trim$(UCase$(CStr(qcSheet.Cells(r, colNilStock).value)))
            If nilStockFlag = "Y" Then
                GoTo SkipRow
            End If
        End If

        ' --- Determine if QC Number is in Production or Inbound ---
        isInProduction = False

        If colFirstReleaseDate > 0 Then
            firstReleaseDate = Trim$(CStr(qcSheet.Cells(r, colFirstReleaseDate).value))
            If firstReleaseDate <> "" Then
                isInProduction = True
            End If
        End If

        If Not isInProduction And colFirstRetestDate > 0 Then
            firstRetestDate = Trim$(CStr(qcSheet.Cells(r, colFirstRetestDate).value))
            If firstRetestDate <> "" Then
                isInProduction = True
            End If
        End If

        ' --- Count based on status and whether it's been counted ---
        If isInProduction Then
            If countedQCs.Exists(qcNumber) Then
                productionCounted = productionCounted + 1
            Else
                productionToCount = productionToCount + 1
            End If
        Else
            If countedQCs.Exists(qcNumber) Then
                inboundCounted = inboundCounted + 1
            Else
                inboundToCount = inboundToCount + 1
            End If
        End If

SkipRow:
    Next r

    ' --- Populate summary cells ---
    With countSheet
        .Cells(8, 3).value = inboundCounted
        .Cells(9, 3).value = inboundToCount
        .Cells(10, 3).value = inboundCounted + inboundToCount

        .Cells(13, 3).value = productionCounted
        .Cells(14, 3).value = productionToCount
        .Cells(15, 3).value = productionCounted + productionToCount
    End With

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateStocktakeSummary"

CleanExit:
    Set countedQCs = Nothing
    Set balanceLookup = Nothing
    Set countSheet = Nothing
    Set qcSheet = Nothing
    Exit Sub
End Sub


