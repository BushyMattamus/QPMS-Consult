Attribute VB_Name = "CSVUtils"
' ========================================================================
' CSV FILE UTILITIES
' ========================================================================
' Handles CSV file operations including reading, writing, parsing,
' and updating QC Number records with change tracking
' ========================================================================

Option Explicit

' Imports a CSV file into the specified worksheet and range using QueryTable.
' Returns True if successful, False otherwise.
Public Function ImportCSVWithQueryTable(ByVal csvPath As String, ByVal ws As Worksheet, ByVal destRange As Range, Optional ByVal clearSheet As Boolean = True) As Boolean
    Dim qt As QueryTable
    On Error GoTo ErrorHandler

    If clearSheet Then
        ws.Cells.Clear
    End If

    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=destRange)
    With qt
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001                 ' UTF-8
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileTrailingMinusNumbers = True
        .PreserveFormatting = True
        .AdjustColumnWidth = False
        .RefreshStyle = xlOverwriteCells
        .TextFileColumnDataTypes = Array(xlTextFormat)
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    Set qt = Nothing
    Set destRange = Nothing
    Set ws = Nothing

    ImportCSVWithQueryTable = True
    Exit Function

ErrorHandler:
    ImportCSVWithQueryTable = False
End Function

Public Function ImportCSVWithOpenText( _
    ByVal csvPath As String, _
    ByVal ws As Worksheet, _
    ByVal destRange As Range, _
    Optional ByVal clearSheet As Boolean = True, _
    Optional ByVal sourceFolderPath As String = "" _
) As Boolean

    On Error GoTo ErrHandler
    
    Dim fso As Object
    Dim tempCsvPath As String
    Dim csvWb As Workbook
    Dim csvWs As Worksheet
    Dim dataRange As Range
    Dim colCount As Long
    Dim fieldInfo As Variant
    Dim i As Long
    Dim ts As Object
    Dim firstLine As String
    Dim commaCount As Long
    Dim sourceFileName As String
    
    ImportCSVWithOpenText = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    sourceFileName = fso.GetFileName(csvPath)
    
    If sourceFolderPath = "" Then
        tempCsvPath = CreateTemporaryFileCopy(sourceFileName)
    Else
        tempCsvPath = CreateTemporaryFileCopy(sourceFileName, sourceFolderPath)
    End If
    
    If tempCsvPath = "" Then
        MsgBox "Failed to create temporary copy of " & sourceFileName, vbCritical, "ImportCSVWithOpenText"
        GoTo CleanExit
    End If
    
    If clearSheet Then
        ws.Cells.Clear
    End If
    
    Set ts = fso.OpenTextFile(tempCsvPath, 1, False)
    If Not ts.AtEndOfStream Then
        firstLine = ts.ReadLine
    End If
    ts.Close
    Set ts = Nothing
    
    commaCount = Len(firstLine) - Len(Replace(firstLine, ",", ""))
    colCount = commaCount + 1
    
    ReDim fieldInfo(1 To colCount)
    For i = 1 To colCount
        fieldInfo(i) = Array(i, xlTextFormat)
    Next i
    
    Workbooks.OpenText _
        fileName:=tempCsvPath, _
        DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Comma:=True, _
        fieldInfo:=fieldInfo
    
    Set csvWb = ActiveWorkbook
    Set csvWs = csvWb.Worksheets(1)
    
    Set dataRange = csvWs.UsedRange
    destRange.Resize(dataRange.Rows.count, dataRange.Columns.count).value = dataRange.value
    
    csvWb.Close SaveChanges:=False
    
    ImportCSVWithOpenText = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ImportCSVWithOpenText"
    On Error Resume Next
    If Not csvWb Is Nothing Then csvWb.Close SaveChanges:=False
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    If Len(tempCsvPath) > 0 Then
        If fso.FileExists(tempCsvPath) Then fso.DeleteFile tempCsvPath
    End If
    Set ts = Nothing
    Set dataRange = Nothing
    Set csvWs = Nothing
    Set csvWb = Nothing
    Set destRange = Nothing
    Set ws = Nothing
    Set fso = Nothing
    Exit Function
End Function

Public Function ReadCSVFile(ByVal filePath As String) As String()
    Const PROC_NAME As String = "ReadCSVFile"
    Dim fileContent As String
    Dim lines() As String
    
    On Error GoTo ErrorHandler
        
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Type = 2
        .Open
        .LoadFromFile filePath
        fileContent = .ReadText
        .Close
    End With
    
    ' Ensure we return a properly dimensioned array
    If Len(Trim$(fileContent)) = 0 Then
        ReDim lines(0 To 0)
        lines(0) = ""
    Else
        lines = Split(ToLF(fileContent), vbLf)
    End If
    
    ReadCSVFile = lines
    GoTo CleanUp
    
ErrorHandler:
    ReDim lines(0 To 0)
    lines(0) = ""
    ReadCSVFile = lines
    
CleanUp:
    Exit Function
End Function

Public Function ParseCSVLine(ByVal line As String) As String()
    Dim elements() As String
    Dim inQuotes As Boolean
    Dim currentElement As String
    Dim i As Long
    Dim char As String
    Dim result() As String
    Dim resultCount As Long
    
    ReDim result(0)
    resultCount = 0
    
    For i = 1 To Len(line)
        char = Mid(line, i, 1)
        
        Select Case char
            Case """"
                inQuotes = Not inQuotes
            Case ","
                If Not inQuotes Then
                    ReDim Preserve result(resultCount)
                    result(resultCount) = currentElement
                    resultCount = resultCount + 1
                    currentElement = ""
                Else
                    currentElement = currentElement & char
                End If
            Case Else
                currentElement = currentElement & char
        End Select
    Next i
    
    ' Add last element
    ReDim Preserve result(resultCount)
    result(resultCount) = currentElement
    
    ParseCSVLine = result
End Function

Private Function ParseCSVLineOptimized(ByVal line As String) As String()
    Dim result() As String
    Dim resultCount As Long
    Dim inQuotes As Boolean
    Dim currentElement As String
    Dim i As Long
    Dim LineLen As Long
    Dim char As String
    Dim estimatedFields As Long
    
    LineLen = Len(line)
    If LineLen = 0 Then
        ReDim result(0)
        result(0) = ""
        ParseCSVLineOptimized = result
        Exit Function
    End If
    
    estimatedFields = 1
    For i = 1 To LineLen
        If Mid$(line, i, 1) = "," Then estimatedFields = estimatedFields + 1
    Next i
    ReDim result(estimatedFields - 1)
    
    resultCount = 0
    currentElement = ""
    inQuotes = False
    
    For i = 1 To LineLen
        char = Mid$(line, i, 1)
        
        Select Case char
            Case """"
                inQuotes = Not inQuotes
            Case ","
                If Not inQuotes Then
                    result(resultCount) = currentElement
                    resultCount = resultCount + 1
                    currentElement = ""
                Else
                    currentElement = currentElement & char
                End If
            Case Else
                currentElement = currentElement & char
        End Select
    Next i
    
    result(resultCount) = currentElement
    
    If resultCount < UBound(result) Then
        ReDim Preserve result(resultCount)
    End If
    
    ParseCSVLineOptimized = result
End Function

Public Function ToLF(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    ToLF = s
End Function

' Helper function to write transactions to RM_Trans.csv with file locking
Public Sub WriteRMTransaction( _
    transType As String, _
    qcNumber As String, _
    qcStat As String, _
    quantity As Double, _
    masterUoM As String)

    
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim transID As String
    
    On Error GoTo WriteRMError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    ' Get next transaction ID (this function handles its own locking)
    transID = GetNextTransactionID()
    
    ' Build CSV line with proper quoting - includes date AND time
    Dim line As String
    line = """" & transID & """," & _
           """" & transType & """," & _
           """" & qcNumber & """," & _
           """" & qcStat & """," & _
           """""," & _
           """""," & _
           """" & CStr(Round(quantity, 3)) & """," & _
           """" & masterUoM & """," & _
           """" & Environ("Username") & """," & _
           """" & GetTransactionTimestamp() & """," & _
           """""," & _
           """""," & _
           """""," & _
           """""," & _
           """"","
    
    ' Append to RM_Trans.csv with exclusive access
    Set ts = fso.OpenTextFile(filePath, 8, True)  ' 8 = ForAppending
    ts.WriteLine line
    ts.Close
    
    GoTo WriteRMCleanup

WriteRMError:
    MsgBox "Error writing to RM_Trans.csv: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close

WriteRMCleanup:
    
    Set ts = Nothing
    Set fso = Nothing
End Sub

' Helper function to write consumption transactions to RM_Trans.csv with all required fields
Public Sub WriteRMConsumptionTransaction( _
    ByVal transType As String, _
    ByVal qcNumber As String, _
    ByVal qcStat As String, _
    ByVal fgBatch As String, _
    ByVal fgBatchStatus As String, _
    ByVal quantity As Double, _
    ByVal masterUoM As String, _
    ByVal adjNote As String, _
    ByVal qcTransID As String, _
    ByVal piTransID As String)
    
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim transID As String
    Dim userName As String
    Dim timestamp As String
    
    On Error GoTo WriteConsumptionError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    ' Get next transaction ID (this function handles its own locking)
    transID = GetNextTransactionID()
    userName = Environ("USERNAME")
    timestamp = GetTransactionTimestamp()
    
    ' Build CSV line with proper quoting - fields ordered to match Constants.GetRMTransHeaders
    Dim line As String
    line = """" & transID & """," & _
            """" & transType & """," & _
            """" & qcNumber & """," & _
            """" & qcStat & """," & _
            """" & fgBatch & """," & _
            """" & fgBatchStatus & """," & _
            """" & CStr(Round(quantity, 3)) & """," & _
            """" & masterUoM & """," & _
            """" & userName & """," & _
            """" & timestamp & """," & _
            """" & qcTransID & """," & _
            """" & piTransID & """," & _
            """" & Replace(adjNote, """", """""") & """," & _
            """""," & _
            """"","
    
    ' Append to RM_Trans.csv with exclusive access
    Set ts = fso.OpenTextFile(filePath, 8, True)  ' 8 = ForAppending
    ts.WriteLine line
    ts.Close
    
    GoTo WriteConsumptionCleanup

WriteConsumptionError:
    MsgBox "Error writing consumption transaction to RM_Trans.csv: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close

WriteConsumptionCleanup:
    Set ts = Nothing
    Set fso = Nothing
End Sub

' --- Update QC Number record with explicit optional parameters ---
Public Function UpdateQCNumberRecord( _
    ByVal qcNumber As String, _
    Optional ByVal standardName As String = "", Optional ByVal rmCode As String = "", _
    Optional ByVal brandName As String = "", Optional ByVal quickCode As String = "", _
    Optional ByVal partARTG As String = "", Optional ByVal vendorName As String = "", _
    Optional ByVal manufacturerName As String = "", Optional ByVal manufacturerLot As String = "", _
    Optional ByVal manufacturerExpiry As String = "", Optional ByVal numberContainers As String = "", _
    Optional ByVal totalQty As String = "", Optional ByVal totalSampleQty As String = "", _
    Optional ByVal masterUoM As String = "", Optional ByVal grDate As String = "", _
    Optional ByVal grBy As String = "", Optional ByVal qcSampleQty As String = "", _
    Optional ByVal retentionSampleQty As String = "", Optional ByVal sampledBy As String = "", _
    Optional ByVal sampleDate As String = "", Optional ByVal qcBalance As String = "", _
    Optional ByVal testingFlag As String = "", Optional ByVal firstReleaseDate As String = "", _
    Optional ByVal firstReleaseBy As String = "", Optional ByVal lastReleaseDate As String = "", _
    Optional ByVal lastReleaseBy As String = "", Optional ByVal firstRetestDate As String = "", _
    Optional ByVal firstRetestBy As String = "", Optional ByVal lastRetestDate As String = "", _
    Optional ByVal lastRetestBy As String = "", Optional ByVal nilStockFlag As String = "", _
    Optional ByVal nilStockFlagDate As String = "", Optional ByVal nilStockFlagBy As String = "", _
    Optional ByVal qcStatusFlag As String = "", Optional ByVal changeReason As String = "", _
    Optional ByVal newRetest As String = "", Optional ByVal qcTested As String = "", _
    Optional ByVal testedDate As String = "", Optional ByVal extendedRetest As String = "", _
    Optional ByVal qcTransID As String = "" _
    ) As Boolean
    
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim fso As Object
    Dim qcFilePath As String
    Dim changesFilePath As String
    Dim tempFilePath As String
    Dim fileContent() As String
    Dim headers() As String
    Dim headerMap As Object
    Dim values() As String
    Dim updatedLine As String
    Dim newQCTransID As String
    Dim foundRecord As Boolean
    Dim i As Long
    Dim tsTemp As Object
    Dim tsChanges As Object
    Dim userName As String
    Dim transDate As String
    Dim changeLogLine As String
    Dim validatedRMCode As String
    Dim validatedMfgExpiry As String
    Dim validatedGRDate As String
    Dim validatedSampleDate As String
    Dim validatedFirstReleaseDate As String
    Dim validatedLastReleaseDate As String
    Dim validatedFirstRetestDate As String
    Dim validatedLastRetestDate As String
    Dim validatedNilStockFlagDate As String
    Dim validatedTestedDate As String
    
    UpdateQCNumberRecord = False
        
    ' --- Step 1: Validate QC Number ---
    If Trim$(qcNumber) = "" Then
        MsgBox "Error: QC Number cannot be empty", vbExclamation, "UpdateQCNumberRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 2: Validate and format RM Code (4-digit padded) ---
    If Trim$(rmCode) <> "" Then
        If Not IsNumeric(rmCode) Then
            MsgBox "Error: RM Code must be numeric", vbExclamation, "UpdateQCNumberRecord"
            GoTo CleanExit
        End If
        validatedRMCode = Format$(CLng(rmCode), "0000")
    End If
    
    ' --- Step 3: Validate date formats (ANSI: YYYY-MM-DD) ---
    If Trim$(manufacturerExpiry) <> "" Then
        validatedMfgExpiry = ValidateAndFormatDate(manufacturerExpiry, "Manufacturer Expiry")
        If validatedMfgExpiry = "" Then GoTo CleanExit
    End If
    
    If Trim$(grDate) <> "" Then
        validatedGRDate = ValidateAndFormatDate(grDate, "GR Date")
        If validatedGRDate = "" Then GoTo CleanExit
    End If
    
    If Trim$(sampleDate) <> "" Then
        validatedSampleDate = ValidateAndFormatDate(sampleDate, "Sample Date")
        If validatedSampleDate = "" Then GoTo CleanExit
    End If
    
    If Trim$(firstReleaseDate) <> "" Then
        validatedFirstReleaseDate = ValidateAndFormatDate(firstReleaseDate, "First Release Date")
        If validatedFirstReleaseDate = "" Then GoTo CleanExit
    End If
    
    If Trim$(lastReleaseDate) <> "" Then
        validatedLastReleaseDate = ValidateAndFormatDate(lastReleaseDate, "Last Release Date")
        If validatedLastReleaseDate = "" Then GoTo CleanExit
    End If
    
    If Trim$(firstRetestDate) <> "" Then
        validatedFirstRetestDate = ValidateAndFormatDate(firstRetestDate, "First Retest Date")
        If validatedFirstRetestDate = "" Then GoTo CleanExit
    End If
    
    If Trim$(lastRetestDate) <> "" Then
        validatedLastRetestDate = ValidateAndFormatDate(lastRetestDate, "Last Retest Date")
        If validatedLastRetestDate = "" Then GoTo CleanExit
    End If
    
    If Trim$(nilStockFlagDate) <> "" And nilStockFlagDate <> "#CLEAR#" Then
        validatedNilStockFlagDate = ValidateAndFormatDate(nilStockFlagDate, "Nil Stock Flag Date")
        If validatedNilStockFlagDate = "" Then GoTo CleanExit
    End If
    
    If Trim$(testedDate) <> "" Then
        validatedTestedDate = ValidateAndFormatDate(testedDate, "Tested Date")
        If validatedTestedDate = "" Then GoTo CleanExit
    End If
        
    ' --- Step 4: Validate flag fields (Y/N only) ---
    If Trim$(qcStatusFlag) <> "" Then
        If Not (UCase$(qcStatusFlag) = "Y" Or UCase$(qcStatusFlag) = "N") Then
            MsgBox "Error: QC Status Flag must be Y or N", vbExclamation, "UpdateQCNumberRecord"
            GoTo CleanExit
        End If
    End If
    
    If Trim$(testingFlag) <> "" Then
        If Not (UCase$(testingFlag) = "Y" Or UCase$(testingFlag) = "N") Then
            MsgBox "Error: Testing Flag must be Y or N", vbExclamation, "UpdateQCNumberRecord"
            GoTo CleanExit
        End If
    End If
    
    If Trim$(nilStockFlag) <> "" Then
        If Not (UCase$(nilStockFlag) = "Y" Or UCase$(nilStockFlag) = "N") Then
            MsgBox "Error: Nil Stock Flag must be Y or N", vbExclamation, "UpdateQCNumberRecord"
            GoTo CleanExit
        End If
    End If
    
    ' --- Step 5: Initialize file paths and objects ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    qcFilePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    changesFilePath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    tempFilePath = qcFilePath & ".tmp"
    userName = Environ("USERNAME")
    transDate = Format$(Date, "YYYY-MM-DD")
    
    ' --- Step 6: Check if file is locked ---
    If IsFileLocked(qcFilePath) Then
        MsgBox "QC_Number.csv is currently locked by another process. Please try again.", vbExclamation, "UpdateQCNumberRecord"
        GoTo CleanExit
    End If
       
    ' --- Step 7: Read original file ---
    fileContent = ReadCSVFile(qcFilePath)
    
    If UBound(fileContent) < 1 Then
        MsgBox "Error: QC_Number.csv is empty or could not be read", vbCritical, "UpdateQCNumberRecord"
        GoTo CleanExit
    End If
        
    ' --- Step 8: Parse headers and create mapping ---
    headers = ParseCSVLine(fileContent(0))
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare
    
    For i = 0 To UBound(headers)
        headerMap(Trim$(Replace(headers(i), """", ""))) = i
    Next i
    
    ' --- Step 9: Get or use provided QC Transaction ID ---
    If Trim$(qcTransID) <> "" Then
        newQCTransID = qcTransID
    Else
        newQCTransID = GetNextQCTransactionID()
    End If

    ' --- Step 10: Build QC Number index and find target record ---
    Dim qcIndex() As String
    Dim targetIndex As Long
    Dim LineLen As Long
    Dim firstCommaPos As Long
    
    ReDim qcIndex(UBound(fileContent))
    targetIndex = -1
        
    For i = 1 To UBound(fileContent)
        LineLen = Len(fileContent(i))
        If LineLen = 0 Then
            qcIndex(i) = ""
        Else
            firstCommaPos = InStr(fileContent(i), ",")
            If firstCommaPos > 0 Then
                qcIndex(i) = Mid$(fileContent(i), 2, firstCommaPos - 3)
            Else
                qcIndex(i) = Mid$(fileContent(i), 2, LineLen - 2)
            End If
            
            If qcIndex(i) = qcNumber And targetIndex = -1 Then
                targetIndex = i
            End If
        End If
    Next i
        
    If targetIndex = -1 Then
        MsgBox "Error: QC Number '" & qcNumber & "' not found in QC_Number.csv", vbExclamation, "UpdateQCNumberRecord"
        GoTo CleanExit
    End If
        
    ' --- Step 11: Parse and update only the target record ---
    values = ParseCSVLineOptimized(fileContent(targetIndex))
    
    If UBound(values) < 0 Then
        MsgBox "Error: Unable to parse record for QC Number '" & qcNumber & "'", vbExclamation, "UpdateQCNumberRecord"
        GoTo CleanExit
    End If
        
    ' --- Step 12: Apply field updates with validated values ---
    If validatedRMCode <> "" Then
        values(headerMap(HEADER_RM_CODE)) = validatedRMCode
    End If
    
    If Trim$(standardName) <> "" Then
        values(headerMap(HEADER_STANDARD_NAME)) = standardName
    End If
    
    If Trim$(quickCode) <> "" Then
        values(headerMap(HEADER_QUICK_CODE)) = quickCode
    End If
    
    If Trim$(brandName) <> "" Then
        values(headerMap(HEADER_BRAND_NAME)) = brandName
    End If
    
    If Trim$(partARTG) <> "" Then
        values(headerMap(HEADER_PART_ARTG)) = partARTG
    End If
    
    If Trim$(vendorName) <> "" Then
        values(headerMap(HEADER_VENDOR_NAME)) = vendorName
    End If
    
    If Trim$(manufacturerName) <> "" Then
        values(headerMap(HEADER_MANUFACTURER_NAME)) = manufacturerName
    End If
    
    If Trim$(manufacturerLot) <> "" Then
        values(headerMap(HEADER_MANUFACTURER_LOT)) = manufacturerLot
    End If
    
    If validatedMfgExpiry <> "" Then
        values(headerMap(HEADER_MANUFACTURER_EXPIRY)) = validatedMfgExpiry
    End If
    
    If Trim$(numberContainers) <> "" Then
        values(headerMap(HEADER_NUMBER_CONTAINERS)) = numberContainers
    End If
    
    If Trim$(totalQty) <> "" Then
        values(headerMap(HEADER_TOTAL_QTY)) = totalQty
    End If
    
    If Trim$(totalSampleQty) <> "" Then
        values(headerMap(HEADER_TOTAL_SAMPLE_QTY)) = totalSampleQty
    End If
    
    If Trim$(masterUoM) <> "" Then
        values(headerMap(HEADER_MASTER_UOM)) = masterUoM
    End If
    
    If validatedGRDate <> "" Then
        values(headerMap(HEADER_GR_DATE)) = validatedGRDate
    End If
    
    If Trim$(grBy) <> "" Then
        values(headerMap(HEADER_GR_BY)) = grBy
    End If
    
    If Trim$(qcSampleQty) <> "" Then
        values(headerMap(HEADER_QC_SAMPLE_QTY)) = qcSampleQty
    End If
    
    If Trim$(retentionSampleQty) <> "" Then
        values(headerMap(HEADER_RETENTION_SAMPLE_QTY)) = retentionSampleQty
    End If
    
    If Trim$(sampledBy) <> "" Then
        values(headerMap(HEADER_SAMPLED_BY)) = sampledBy
    End If
    
    If validatedSampleDate <> "" Then
        values(headerMap(HEADER_SAMPLE_DATE)) = validatedSampleDate
    End If

    If qcBalance <> "" Then
        values(headerMap(HEADER_QC_BALANCE)) = qcBalance
    End If
    
    If Trim$(testingFlag) <> "" Then
        values(headerMap(HEADER_TESTING_FLAG)) = UCase$(testingFlag)
    End If
    
    If validatedFirstReleaseDate <> "" Then
        values(headerMap(HEADER_FIRST_RELEASE_DATE)) = validatedFirstReleaseDate
    End If
    
    If Trim$(firstReleaseBy) <> "" Then
        values(headerMap(HEADER_FIRST_RELEASE_BY)) = firstReleaseBy
    End If
    
    If validatedLastReleaseDate <> "" Then
        values(headerMap(HEADER_LAST_RELEASE_DATE)) = validatedLastReleaseDate
    End If
    
    If Trim$(lastReleaseBy) <> "" Then
        values(headerMap(HEADER_LAST_RELEASE_BY)) = lastReleaseBy
    End If
    
    If validatedFirstRetestDate <> "" Then
        values(headerMap(HEADER_FIRST_RETEST_DATE)) = validatedFirstRetestDate
    End If
    
    If Trim$(firstRetestBy) <> "" Then
        values(headerMap(HEADER_FIRST_RETEST_BY)) = firstRetestBy
    End If
    
    If validatedLastRetestDate <> "" Then
        values(headerMap(HEADER_LAST_RETEST_DATE)) = validatedLastRetestDate
    End If
    
    If Trim$(lastRetestBy) <> "" Then
        values(headerMap(HEADER_LAST_RETEST_BY)) = lastRetestBy
    End If
    
    If Trim$(nilStockFlag) <> "" Then
        values(headerMap(HEADER_NIL_STOCK_FLAG)) = UCase$(nilStockFlag)
    End If
    
     If Trim$(nilStockFlagDate) <> "" Then
        If nilStockFlagDate = "#CLEAR#" Then
            values(headerMap(HEADER_NIL_STOCK_FLAG_DATE)) = ""
        Else
            values(headerMap(HEADER_NIL_STOCK_FLAG_DATE)) = validatedNilStockFlagDate
        End If
    End If
    
    If Trim$(nilStockFlagBy) <> "" Then
        If nilStockFlagBy = "#CLEAR#" Then
            values(headerMap(HEADER_NIL_STOCK_FLAG_BY)) = ""
        Else
            values(headerMap(HEADER_NIL_STOCK_FLAG_BY)) = nilStockFlagBy
        End If
    End If
    
    If Trim$(qcStatusFlag) <> "" Then
        values(headerMap(HEADER_QC_STATUS_FLAG)) = UCase$(qcStatusFlag)
    End If

    If Trim$(changeReason) <> "" Then
        values(headerMap(HEADER_CHANGE_REASON)) = changeReason
    End If
    
    If Trim$(newRetest) <> "" Then
        values(headerMap(HEADER_NEW_RETEST)) = newRetest
    End If

    If Trim$(transDate) <> "" Then
        values(headerMap(HEADER_TRANS_DATE)) = transDate
    End If

    If Trim$(userName) <> "" Then
        values(headerMap(HEADER_TRANS_USER)) = userName
    End If
    
    If Trim$(qcTested) <> "" Then
        values(headerMap(HEADER_QC_TESTED)) = qcTested
    End If
    
    If validatedTestedDate <> "" Then
        values(headerMap(HEADER_TESTED_DATE)) = validatedTestedDate
    End If
    
    If Trim$(extendedRetest) <> "" Then
        values(headerMap(HEADER_EXTENDED_RETEST)) = extendedRetest
    End If

    If Trim$(newQCTransID) <> "" Then
        values(headerMap(HEADER_QC_TRANS_ID)) = newQCTransID
    End If
    
    updatedLine = BuildCSVLine(values)
        
    ' --- Step 13: Build temp file content in memory ---
    Dim tempContent As String
    tempContent = fileContent(0) & vbCrLf
    
    For i = 1 To UBound(fileContent)
        If i = targetIndex Then
            tempContent = tempContent & updatedLine
        Else
            tempContent = tempContent & fileContent(i)
        End If
        
        If i < UBound(fileContent) Then
            tempContent = tempContent & vbCrLf
        End If
    Next i
    
    ' --- Step 14: Write temp file in single operation ---
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Type = 2
        .Open
        .WriteText tempContent, 0
        .SaveToFile tempFilePath, 2
        .Close
    End With
    
    ' --- Step 15: Append to QC_Number_Changes.csv ---
    Set tsChanges = fso.OpenTextFile(changesFilePath, 8, True)
    tsChanges.WriteLine updatedLine
    tsChanges.Close
    Set tsChanges = Nothing
    
    ' --- Step 16: Replace original file ---
    fso.DeleteFile qcFilePath
    fso.MoveFile tempFilePath, qcFilePath

    UpdateQCNumberRecord = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "UpdateQCNumberRecord"
    On Error Resume Next
    If Not tsTemp Is Nothing Then tsTemp.Close
    If Not tsChanges Is Nothing Then tsChanges.Close
    If fso.FileExists(tempFilePath) Then fso.DeleteFile tempFilePath
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    If Not tsTemp Is Nothing Then Set tsTemp = Nothing
    If Not tsChanges Is Nothing Then Set tsChanges = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    Set headerMap = Nothing
    Exit Function
    
End Function

' --- Validate and format date to ANSI (YYYY-MM-DD) ---
Private Function ValidateAndFormatDate(ByVal dateStr As String, ByVal fieldName As String) As String
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim testDate As Date
    
    ValidateAndFormatDate = ""
    
    ' --- Step 1: Attempt date conversion ---
    testDate = CDate(dateStr)
    
    ' --- Step 2: Format to ANSI ---
    ValidateAndFormatDate = Format$(testDate, "YYYY-MM-DD")
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error: " & fieldName & " is not a valid date format", vbExclamation, "ValidateAndFormatDate"
    ValidateAndFormatDate = ""
    
CleanExit:
    Exit Function
End Function

' --- Build CSV line from array of values ---
Private Function BuildCSVLine(ByRef values() As String) As String
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim i As Long
    Dim result As String
    
    ' --- Step 1: Build quoted CSV line ---
    For i = 0 To UBound(values)
        If i > 0 Then result = result & ","
        result = result & """" & Replace(values(i), """", """""") & """"
    Next i
    
    BuildCSVLine = result
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "BuildCSVLine"
    BuildCSVLine = ""
    
CleanExit:
    Exit Function
End Function
