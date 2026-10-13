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
    destRange.Resize(dataRange.Rows.Count, dataRange.Columns.Count).value = dataRange.value
    
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

Public Function ExportCSVToExcel(ByVal csvName As String, ByVal exportPrefix As String) As Boolean
    Const PROC_NAME As String = "ExportCSVToExcel"
    Dim srcPath As String
    Dim destPath As String
    Dim wb As Workbook
    
    On Error GoTo Cleanup
    
    srcPath = BuildFilePath(GetWorkbookPath(), FOLDER_TABLES, csvName)
    destPath = BuildExportPath(exportPrefix & GenerateTimestamp() & ".xlsx")
    
    EnsureFolderExists BuildFilePath(GetWorkbookPath(), FOLDER_EXPORTS)
    
    DisableAppUpdates
    Set wb = Workbooks.Open(fileName:=srcPath, ReadOnly:=True)
    
    wb.SaveAs fileName:=destPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    
    MsgBox "Export completed:" & vbCrLf & destPath, vbInformation
    ExportCSVToExcel = True
    GoTo Cleanup
    
Cleanup:
    RestoreAppState
    Exit Function
End Function

Public Function ReadCSVFile(ByVal filePath As String) As String()
    Const PROC_NAME As String = "ReadCSVFile"
    Dim fileContent As String
    Dim lines() As String
    Dim fso As Object
    Dim ts As Object
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close
    
    ' Ensure we return a properly dimensioned array
    If Len(Trim$(fileContent)) = 0 Then
        ReDim lines(0 To 0)
        lines(0) = ""
    Else
        lines = Split(ToLF(fileContent), vbLf)
    End If
    
    ReadCSVFile = lines
    GoTo Cleanup
    
ErrorHandler:
    ReDim lines(0 To 0)
    lines(0) = ""
    ReadCSVFile = lines
    
Cleanup:
    If Not ts Is Nothing Then Set ts = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
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
    Dim lineLen As Long
    Dim char As String
    Dim estimatedFields As Long
    
    lineLen = Len(line)
    If lineLen = 0 Then
        ReDim result(0)
        result(0) = ""
        ParseCSVLineOptimized = result
        Exit Function
    End If
    
    estimatedFields = 1
    For i = 1 To lineLen
        If Mid$(line, i, 1) = "," Then estimatedFields = estimatedFields + 1
    Next i
    ReDim result(estimatedFields - 1)
    
    resultCount = 0
    currentElement = ""
    inQuotes = False
    
    For i = 1 To lineLen
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

' --- Update Finished Goods Batch record with optional field parameters ---
Public Function UpdateFinishedGoodsBatchRecord( _
    ByVal batchNumber As String, _
    Optional ByVal pmiNumber As String = "", _
    Optional ByVal pmiVersion As String = "", _
    Optional ByVal productDescription As String = "", _
    Optional ByVal shelfLifeNo As String = "", _
    Optional ByVal shelfLife As String = "", _
    Optional ByVal manufacturingDate As String = "", _
    Optional ByVal expiryDate As String = "", _
    Optional ByVal dateCreated As String = "", _
    Optional ByVal timeCreated As String = "", _
    Optional ByVal batchStatus As String = "", _
    Optional ByVal createdBy As String = "", _
    Optional ByVal dateUpdated As String = "", _
    Optional ByVal updatedBy As String = "", _
    Optional ByVal updateComments As String = "", _
    Optional ByVal auditTransID As String = "", _
    Optional ByVal scanned As String = "" _
    ) As Boolean
    
    ' --- Error handling setup ---
    On Error GoTo ErrHandler
    
    ' --- Variable declarations ---
    Dim fso As Object
    Dim filePath As String
    Dim tempFilePath As String
    Dim fileContent() As String
    Dim headers() As String
    Dim headerMap As Object
    Dim values() As String
    Dim updatedLine As String
    Dim foundRecord As Boolean
    Dim i As Long
    Dim targetIndex As Long
    Dim batchIndex() As String
    Dim lineLen As Long
    Dim firstCommaPos As Long
    Dim tsTemp As Object
    Dim tempContent As String
    
    UpdateFinishedGoodsBatchRecord = False
    
    ' --- Step 1: Validate Batch Number ---
    If Trim$(batchNumber) = "" Then
        MsgBox "Error: Batch Number cannot be empty", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 2: Initialize file paths and objects ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_FINISHED_GOODS_BATCH_MASTER)
    tempFilePath = filePath & ".tmp"
    
    ' --- Step 3: Check if file is locked ---
    If IsFileLocked(filePath) Then
        MsgBox "FinishedGoodsBatchNumberMasterFile.csv is currently locked by another process. Please try again.", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 4: Read original file ---
    fileContent = ReadCSVFile(filePath)
    
    If UBound(fileContent) < 1 Then
        MsgBox "Error: FinishedGoodsBatchNumberMasterFile.csv is empty or could not be read", vbCritical, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 5: Parse headers and create mapping ---
    headers = ParseCSVLine(fileContent(0))
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare
    
    For i = 0 To UBound(headers)
        headerMap(Trim$(Replace(headers(i), """", ""))) = i
    Next i
    
    ' --- Step 6: Build Batch Number index and find target record ---
    ReDim batchIndex(UBound(fileContent))
    targetIndex = -1
        
    For i = 1 To UBound(fileContent)
        lineLen = Len(fileContent(i))
        If lineLen = 0 Then
            batchIndex(i) = ""
        Else
            firstCommaPos = InStr(fileContent(i), ",")
            If firstCommaPos > 0 Then
                batchIndex(i) = Mid$(fileContent(i), 2, firstCommaPos - 3)
            Else
                batchIndex(i) = Mid$(fileContent(i), 2, lineLen - 2)
            End If
                        
            If batchIndex(i) = batchNumber And targetIndex = -1 Then
                targetIndex = i
            End If
        End If
    Next i
        
    If targetIndex = -1 Then
        MsgBox "Error: Batch Number '" & batchNumber & "' not found in FinishedGoodsBatchNumberMasterFile.csv", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 7: Parse and update only the target record ---
    values = ParseCSVLineOptimized(fileContent(targetIndex))
    
    If UBound(values) < 0 Then
        MsgBox "Error: Unable to parse record for Batch Number '" & batchNumber & "'", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    ' --- Step 8: Apply field updates ---
    If Trim$(pmiNumber) <> "" Then
        values(headerMap(FGBNM_HEADER_PMI_NUMBER)) = pmiNumber
    End If
    
    If Trim$(pmiVersion) <> "" Then
        values(headerMap(FGBNM_HEADER_PMI_VERSION)) = pmiVersion
    End If
    
    If Trim$(productDescription) <> "" Then
        values(headerMap(FGBNM_HEADER_PRODUCT_DESCRIPTION)) = productDescription
    End If
    
    If Trim$(shelfLifeNo) <> "" Then
        values(headerMap(FGBNM_HEADER_SHELF_LIFE_NO)) = shelfLifeNo
    End If
    
    If Trim$(shelfLife) <> "" Then
        values(headerMap(FGBNM_HEADER_SHELF_LIFE)) = shelfLife
    End If
    
    If Trim$(manufacturingDate) <> "" Then
        values(headerMap(FGBNM_HEADER_MANUFACTURING_DATE)) = manufacturingDate
    End If
    
    If Trim$(expiryDate) <> "" Then
        values(headerMap(FGBNM_HEADER_EXPIRY_DATE)) = expiryDate
    End If
    
    If Trim$(dateCreated) <> "" Then
        values(headerMap(FGBNM_HEADER_DATE_CREATED)) = dateCreated
    End If
    
    If Trim$(timeCreated) <> "" Then
        values(headerMap(FGBNM_HEADER_TIME_CREATED)) = timeCreated
    End If
    
    If Trim$(batchStatus) <> "" Then
        values(headerMap(FGBNM_HEADER_BATCH_STATUS)) = batchStatus
    End If
    
    If Trim$(createdBy) <> "" Then
        values(headerMap(FGBNM_HEADER_CREATED_BY)) = createdBy
    End If
    
    values(headerMap(FGBNM_HEADER_DATE_UPDATED)) = GetTransactionTimestamp()
    
    If Trim$(updatedBy) <> "" Then
        values(headerMap(FGBNM_HEADER_UPDATED_BY)) = updatedBy
    End If
    
    If Trim$(updateComments) <> "" Then
        values(headerMap(FGBNM_HEADER_UPDATE_COMMENTS)) = updateComments
    End If
    
    If Trim$(scanned) <> "" Then
        values(headerMap(FGBNM_HEADER_SCANNED)) = scanned
    End If
    
    updatedLine = BuildCSVLine(values)
    
    ' --- Step 9: Build temp file content in memory ---
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
    
    ' --- Step 10: Write temp file (UTF-8 without BOM) ---
    Set tsTemp = fso.OpenTextFile(tempFilePath, 2, True)
    tsTemp.Write tempContent
    tsTemp.Close
    
    ' --- Step 11: Replace original file ---
    fso.DeleteFile filePath
    fso.MoveFile tempFilePath, filePath
    
    ' --- Step 12: Get next Trans_Id and build audit line ---
    Dim auditFilePath As String
    Dim tsAudit As Object
    Dim finalTransID As String
    Dim auditLine As String
    
    If auditTransID <> "" Then
        finalTransID = auditTransID
    Else
        finalTransID = GetNextFinishedGoodsBatchTransID()
    End If
    
    auditLine = updatedLine & ",""" & finalTransID & """"
    
    auditFilePath = GetWorkbookTableFilePath(FILE_FINISHED_GOODS_BATCH_AUDIT)
    Set tsAudit = fso.OpenTextFile(auditFilePath, 8, True)  ' 8 = ForAppending
    tsAudit.WriteLine auditLine
    tsAudit.Close
    Set tsAudit = Nothing
    
    UpdateFinishedGoodsBatchRecord = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "UpdateFinishedGoodsBatchRecord"
    On Error Resume Next
    If fso.FileExists(tempFilePath) Then fso.DeleteFile tempFilePath
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    If Not tsTemp Is Nothing Then Set tsTemp = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    Set headerMap = Nothing
    Exit Function
End Function

Public Sub WriteProductionOutputRecord(ByRef outputData() As Variant)
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim outTransID As String
    Dim userName As String
    Dim timestamp As String
    
    On Error GoTo WriteOutputError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_PRODUCTION_OUTPUT)
    
    outTransID = GetNextOutputRecordNumber()
    userName = Environ("USERNAME")
    timestamp = Format$(Now, "dd/mm/yyyy hh:mm:ss")
    
    Dim line As String
    Dim i As Long
    
    line = """" & outTransID & """"
    For i = 1 To UBound(outputData)
        line = line & ",""" & Replace(CStr(outputData(i)), """", """""") & """"
    Next i
    
    Set ts = fso.OpenTextFile(filePath, 8, True)
    ts.WriteLine line
    ts.Close
    
    GoTo WriteOutputCleanup

WriteOutputError:
    MsgBox "Error writing to Production_Output.csv: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close

WriteOutputCleanup:
    Set ts = Nothing
    Set fso = Nothing
End Sub
