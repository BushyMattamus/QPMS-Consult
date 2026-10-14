Attribute VB_Name = "CSVUtils"
' ========================================================================
' CSV FILE UTILITIES
' ========================================================================
' Handles CSV file operations including reading, writing, parsing,
' and updating QC Number records with change tracking
' ========================================================================

Option Explicit

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
    
    srcPath = GetWorkbookTableFilePath(csvName)
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
    Dim timeStamp As String
    
    On Error GoTo WriteConsumptionError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    ' Get next transaction ID (this function handles its own locking)
    transID = GetNextTransactionID()
    userName = Environ("USERNAME")
    timeStamp = GetTransactionTimestamp()
    
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
            """" & timeStamp & """," & _
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
    Optional ByVal scanned As String = "" _
    ) As Boolean
    
    On Error GoTo ErrHandler
    
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
    Dim currentBatchStatus As String
    Dim requestedBatchStatus As String
    Dim finalBatchStatus As String
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
    
    currentBatchStatus = Trim$(Replace(values(headerMap(FGBNM_HEADER_BATCH_STATUS)), """", ""))
    
    If UCase$(currentBatchStatus) = UCase$(FGB_STATUS_REJECTED) Or UCase$(currentBatchStatus) = UCase$(FGB_STATUS_CANCELLED) Then
        MsgBox "Error: Cannot update batch with status '" & currentBatchStatus & "'." & vbCrLf & _
               "Rejected and Cancelled batches cannot be modified.", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    requestedBatchStatus = Trim$(batchStatus)
    
    If requestedBatchStatus <> "" Then
        finalBatchStatus = ResolveProgressiveBatchStatus(currentBatchStatus, requestedBatchStatus)
    Else
        finalBatchStatus = currentBatchStatus
    End If
    
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
    
    values(headerMap(FGBNM_HEADER_BATCH_STATUS)) = finalBatchStatus
    
    If Trim$(createdBy) <> "" Then
        values(headerMap(FGBNM_HEADER_CREATED_BY)) = createdBy
    End If

    values(headerMap(FGBNM_HEADER_DATE_UPDATED)) = GetTransactionTimestamp()
    
    If Trim$(updatedBy) <> "" Then
        values(headerMap(FGBNM_HEADER_UPDATED_BY)) = updatedBy
    End If
    
    If Trim$(updateComments) <> "" Then
        values(headerMap(FGBNM_HEADER_UPDATE_COMMENTS)) = updateComments
    Else
        values(headerMap(FGBNM_HEADER_UPDATE_COMMENTS)) = ""
    End If
    
    If Trim$(scanned) <> "" Then
        values(headerMap(FGBNM_HEADER_SCANNED)) = scanned
    End If
    
    updatedLine = BuildCSVLine(values)
    
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

    Set tsTemp = fso.OpenTextFile(tempFilePath, 2, True)
    tsTemp.Write tempContent
    tsTemp.Close
    
    ' --- Step 11: Replace original file ---
    fso.DeleteFile filePath
    fso.MoveFile tempFilePath, filePath
    
    ' --- Step 12: Get next Trans_Id and build audit line ---
    Dim auditFilePath As String
    Dim tsAudit As Object
    Dim auditTransID As String
    Dim auditLine As String
    
    auditTransID = GetNextFinishedGoodsBatchTransID()
    auditLine = updatedLine & ",""" & auditTransID & """"
    
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

' --- Append new Label Print record to Label_Prints.csv ---
Public Function AppendLabelPrintRecord( _
    ByVal batchNumber As String, _
    ByVal quantity As String, _
    ByVal pmiNumber As String, _
    ByVal expiryDate As String, _
    Optional ByVal qidCode As String = "", _
    Optional ByVal mfgLotNo As String = "", _
    Optional ByVal mfgLotNo2 As String = "", _
    Optional ByVal averyLabelCode As String = "", _
    Optional ByVal labelsPerSheet As String = "", _
    Optional ByVal productID As String = "", _
    Optional ByVal cpmuRef As String = "", _
    Optional ByVal productDescription As String = "", _
    Optional ByVal printType As String = "", _
    Optional ByVal reprintReason As String = "", _
    Optional ByVal originalPrintID As String = "", _
    Optional ByVal comments As String = "", _
    Optional ByVal reprintReasonType As String = "" _
    ) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Dim filePath As String
    Dim tempFilePath As String
    Dim tsAppend As Object
    Dim values(0 To 22) As String
    Dim csvLine As String
    Dim userName As String
    Dim currentDateTime As String
    Dim printID As String
    Dim quantityNum As Double
    Dim labelsPerSheetNum As Double
    
    AppendLabelPrintRecord = False
    
    If Trim$(batchNumber) = "" Then
        MsgBox "Error: Batch Number cannot be empty", vbExclamation, "AppendLabelPrintRecord"
        GoTo CleanExit
    End If
    
    If Trim$(quantity) = "" Or Not IsNumeric(quantity) Then
        MsgBox "Error: Quantity must be a valid number", vbExclamation, "AppendLabelPrintRecord"
        GoTo CleanExit
    End If
    
    quantityNum = CDbl(quantity)
    If quantityNum <= 0 Then
        MsgBox "Error: Quantity must be greater than zero", vbExclamation, "AppendLabelPrintRecord"
        GoTo CleanExit
    End If
    
    If Trim$(pmiNumber) = "" Then
        MsgBox "Error: PMI Number cannot be empty", vbExclamation, "AppendLabelPrintRecord"
        GoTo CleanExit
    End If
    
    If Trim$(expiryDate) = "" Then
        MsgBox "Error: Expiry Date cannot be empty", vbExclamation, "AppendLabelPrintRecord"
        GoTo CleanExit
    End If
    
    If Trim$(labelsPerSheet) <> "" Then
        If Not IsNumeric(labelsPerSheet) Then
            MsgBox "Error: Labels_per_sheet must be a valid number", vbExclamation, "AppendLabelPrintRecord"
            GoTo CleanExit
        End If
        labelsPerSheetNum = CDbl(labelsPerSheet)
        If labelsPerSheetNum <= 0 Then
            MsgBox "Error: Labels_per_sheet must be greater than zero", vbExclamation, "AppendLabelPrintRecord"
            GoTo CleanExit
        End If
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_LABEL_DETAILS)
    
    If Not fso.FileExists(filePath) Then
        MsgBox "Error: Label_Prints.csv file not found at: " & filePath, vbCritical, "AppendLabelPrintRecord"
        GoTo CleanExit
    End If
    
    If IsFileLocked(filePath) Then
        MsgBox "Label_Prints.csv is currently locked by another process. Please try again.", vbExclamation, "AppendLabelPrintRecord"
        GoTo CleanExit
    End If
    
    userName = Environ("USERNAME")
    currentDateTime = GetTransactionTimestamp()
    printID = GetNextLabelPrintID()
    
    values(0) = qidCode
    values(1) = batchNumber
    values(2) = mfgLotNo
    values(3) = mfgLotNo2
    values(4) = expiryDate
    values(5) = quantity
    values(6) = userName
    values(7) = currentDateTime
    values(8) = averyLabelCode
    values(9) = labelsPerSheet
    values(10) = quantity
    values(11) = productID
    values(12) = cpmuRef
    values(13) = productDescription
    values(14) = printType
    values(15) = reprintReason
    values(16) = "X"
    values(17) = currentDateTime
    values(18) = printID
    values(19) = originalPrintID
    values(20) = comments
    values(21) = reprintReasonType
    values(22) = userName
    
    csvLine = BuildCSVLine(values)
    
    Set tsAppend = fso.OpenTextFile(filePath, 8, False)
    tsAppend.WriteLine csvLine
    tsAppend.Close
    
    AppendLabelPrintRecord = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "AppendLabelPrintRecord"
    On Error Resume Next
    If Not tsAppend Is Nothing Then
        tsAppend.Close
    End If
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    If Not tsAppend Is Nothing Then Set tsAppend = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' RESOLVE PROGRESSIVE BATCH STATUS
' ========================================================================
' Ensures batch status only progresses forward through the workflow
' Sequence: Created --> Printed --> InProgress --> Complete
' Returns the more advanced status between current and requested
' ========================================================================

Private Function ResolveProgressiveBatchStatus(ByVal currentStatus As String, ByVal requestedStatus As String) As String
    On Error GoTo ErrHandler
    
    Dim statusHierarchy As Object
    Dim currentLevel As Long
    Dim requestedLevel As Long
    Dim debugMsg As String
    
    Set statusHierarchy = CreateObject("Scripting.Dictionary")
    statusHierarchy.CompareMode = vbTextCompare
    
    statusHierarchy(FGB_STATUS_CREATED) = 1
    statusHierarchy(FGB_STATUS_PRINTED) = 2
    statusHierarchy(FGB_STATUS_INPROGRESS) = 3
    statusHierarchy(FGB_STATUS_COMPLETE) = 4
    
    If Not statusHierarchy.Exists(currentStatus) Then
        currentLevel = 0
    Else
        currentLevel = statusHierarchy(currentStatus)
    End If
    
    If Not statusHierarchy.Exists(requestedStatus) Then
        requestedLevel = 0
    Else
        requestedLevel = statusHierarchy(requestedStatus)
    End If
    
    If requestedLevel > currentLevel Then
        ResolveProgressiveBatchStatus = requestedStatus
    Else
        ResolveProgressiveBatchStatus = currentStatus
    End If
    
    GoTo CleanExit
    
ErrHandler:
    ResolveProgressiveBatchStatus = currentStatus
    GoTo CleanExit
    
CleanExit:
    Set statusHierarchy = Nothing
End Function

' ========================================================================
' LABEL REGISTER RECORD EXISTS CHECK
' ========================================================================
' Checks if a PMI Number already exists in Label_Register.csv
' Returns True if found, False otherwise
' Uses temporary file copy and ADO for fast, safe operation
' ========================================================================

Public Function LabelRegisterRecordExists(ByVal pmiNumber As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    
    LabelRegisterRecordExists = False
    
    If Trim$(pmiNumber) = "" Then
        GoTo CleanExit
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    tempPath = CreateTemporaryFileCopy(FILE_LABEL_REGISTER)
    
    If tempPath = "" Then
        GoTo CleanExit
    End If
    
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    Set conn = OpenADOConnection(folderPath)
    
    If conn Is Nothing Then
        GoTo CleanExit
    End If
    
    sql = "SELECT COUNT(*) FROM [" & fileName & "] " & _
          "WHERE [" & LR_HEADER_PMI_NUMBER & "] = '" & Replace(pmiNumber, "'", "''") & "'"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        LabelRegisterRecordExists = (rs.Fields(0).value > 0)
    End If
    
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LabelRegisterRecordExists"
    LabelRegisterRecordExists = False
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' UPDATE LABEL REGISTER RECORD
' ========================================================================
' Updates or deletes an existing Label_Register record by PMI_Number
' If deleteRecord = True, removes the record from the file
' If deleteRecord = False, updates the specified fields
' ========================================================================

Public Function UpdateLabelRegisterRecord( _
    ByVal pmiNumber As String, _
    Optional ByVal deleteRecord As Boolean = False, _
    Optional ByVal averyCode As String = "", _
    Optional ByVal specialInstructions As String = "", _
    Optional ByVal labelsPerBatch As String = "" _
    ) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Dim filePath As String
    Dim tempFilePath As String
    Dim fileContent() As String
    Dim headers() As String
    Dim headerMap As Object
    Dim values() As String
    Dim updatedLine As String
    Dim i As Long
    Dim targetIndex As Long
    Dim pmiIndex() As String
    Dim lineLen As Long
    Dim firstCommaPos As Long
    Dim tempContent As String
    Dim tsTemp As Object
    Dim userName As String
    Dim currentDateTime As String
    Dim labelsPerBatchNum As Double
    
    UpdateLabelRegisterRecord = False
    
    If Trim$(pmiNumber) = "" Then
        MsgBox "Error: PMI Number cannot be empty", vbExclamation, "UpdateLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    If Not deleteRecord Then
        If Trim$(labelsPerBatch) <> "" Then
            If Not IsNumeric(labelsPerBatch) Then
                MsgBox "Error: Labels_Per_Batch must be a valid number", vbExclamation, "UpdateLabelRegisterRecord"
                GoTo CleanExit
            End If
            labelsPerBatchNum = CDbl(labelsPerBatch)
            If labelsPerBatchNum <= 0 Then
                MsgBox "Error: Labels_Per_Batch must be greater than zero", vbExclamation, "UpdateLabelRegisterRecord"
                GoTo CleanExit
            End If
        End If
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_LABEL_REGISTER)
    tempFilePath = filePath & ".tmp"
    
    If IsFileLocked(filePath) Then
        MsgBox "Label_Register.csv is currently locked by another process. Please try again.", vbExclamation, "UpdateLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    fileContent = ReadCSVFile(filePath)
    
    If UBound(fileContent) < 1 Then
        MsgBox "Error: Label_Register.csv is empty or could not be read", vbCritical, "UpdateLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    headers = ParseCSVLine(fileContent(0))
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare
    
    For i = 0 To UBound(headers)
        headerMap(Trim$(Replace(headers(i), """", ""))) = i
    Next i
    
    ReDim pmiIndex(UBound(fileContent))
    targetIndex = -1
    
    For i = 1 To UBound(fileContent)
        lineLen = Len(fileContent(i))
        If lineLen = 0 Then
            pmiIndex(i) = ""
        Else
            firstCommaPos = InStr(fileContent(i), ",")
            If firstCommaPos > 0 Then
                pmiIndex(i) = Mid$(fileContent(i), 2, firstCommaPos - 3)
            Else
                pmiIndex(i) = Mid$(fileContent(i), 2, lineLen - 2)
            End If
            
            If pmiIndex(i) = pmiNumber And targetIndex = -1 Then
                targetIndex = i
            End If
        End If
    Next i
    
    If targetIndex = -1 Then
        MsgBox "Error: PMI Number '" & pmiNumber & "' not found in Label_Register.csv", vbExclamation, "UpdateLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    If deleteRecord Then
        tempContent = fileContent(0) & vbCrLf
        
        For i = 1 To UBound(fileContent)
            If i <> targetIndex And Len(Trim$(fileContent(i))) > 0 Then
                tempContent = tempContent & fileContent(i)
                If i < UBound(fileContent) Then
                    tempContent = tempContent & vbCrLf
                End If
            End If
        Next i
    Else
        values = ParseCSVLineOptimized(fileContent(targetIndex))
        
        If UBound(values) < 0 Then
            MsgBox "Error: Unable to parse record for PMI Number '" & pmiNumber & "'", vbExclamation, "UpdateLabelRegisterRecord"
            GoTo CleanExit
        End If
        
        values(headerMap(LR_HEADER_AVERY_CODE)) = averyCode
        
        values(headerMap(LR_HEADER_SPECIAL_INSTRUCTIONS)) = specialInstructions
        
        values(headerMap(LR_HEADER_LABELS_PER_BATCH)) = labelsPerBatch
        
        userName = Environ("USERNAME")
        currentDateTime = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        
        values(headerMap(LR_HEADER_UPDATED_BY)) = userName
        values(headerMap(LR_HEADER_LAST_CHANGE_DATE)) = currentDateTime
        
        updatedLine = BuildCSVLine(values)
        
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
    End If
    
    Set tsTemp = fso.OpenTextFile(tempFilePath, 2, True)
    tsTemp.Write tempContent
    tsTemp.Close
    
    fso.DeleteFile filePath
    fso.MoveFile tempFilePath, filePath
    
    UpdateLabelRegisterRecord = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "UpdateLabelRegisterRecord"
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(tempFilePath) Then fso.DeleteFile tempFilePath
    End If
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    Set tsTemp = Nothing
    Set headerMap = Nothing
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' APPEND LABEL REGISTER RECORD
' ========================================================================
' Adds a new record to Label_Register.csv
' Validates that PMI_Number is unique before appending
' ========================================================================

Public Function AppendLabelRegisterRecord( _
    ByVal pmiNumber As String, _
    ByVal averyCode As String, _
    Optional ByVal specialInstructions As String = "", _
    Optional ByVal labelsPerBatch As String = "" _
    ) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Dim filePath As String
    Dim tsAppend As Object
    Dim values(0 To 5) As String
    Dim csvLine As String
    Dim userName As String
    Dim currentDateTime As String
    Dim labelsPerBatchNum As Double
    
    AppendLabelRegisterRecord = False
    
    If Trim$(pmiNumber) = "" Then
        MsgBox "Error: PMI Number cannot be empty", vbExclamation, "AppendLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    If Trim$(averyCode) = "" Then
        MsgBox "Error: Avery Code cannot be empty", vbExclamation, "AppendLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    If Trim$(labelsPerBatch) <> "" Then
        If Not IsNumeric(labelsPerBatch) Then
            MsgBox "Error: Labels_Per_Batch must be a valid number", vbExclamation, "AppendLabelRegisterRecord"
            GoTo CleanExit
        End If
        labelsPerBatchNum = CDbl(labelsPerBatch)
        If labelsPerBatchNum <= 0 Then
            MsgBox "Error: Labels_Per_Batch must be greater than zero", vbExclamation, "AppendLabelRegisterRecord"
            GoTo CleanExit
        End If
    End If
    
    If LabelRegisterRecordExists(pmiNumber) Then
        MsgBox "Error: PMI Number '" & pmiNumber & "' already exists in Label_Register.csv", vbExclamation, "AppendLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_LABEL_REGISTER)
    
    If Not fso.FileExists(filePath) Then
        MsgBox "Error: Label_Register.csv file not found at: " & filePath, vbCritical, "AppendLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    If IsFileLocked(filePath) Then
        MsgBox "Label_Register.csv is currently locked by another process. Please try again.", vbExclamation, "AppendLabelRegisterRecord"
        GoTo CleanExit
    End If
    
    userName = Environ("USERNAME")
    currentDateTime = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    
    values(0) = pmiNumber
    values(1) = averyCode
    values(2) = userName
    values(3) = currentDateTime
    values(4) = specialInstructions
    values(5) = labelsPerBatch
    
    csvLine = BuildCSVLine(values)
    
    Set tsAppend = fso.OpenTextFile(filePath, 8, False)
    tsAppend.WriteLine csvLine
    tsAppend.Close
    
    AppendLabelRegisterRecord = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "AppendLabelRegisterRecord"
    On Error Resume Next
    If Not tsAppend Is Nothing Then
        tsAppend.Close
    End If
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    Set tsAppend = Nothing
    Set fso = Nothing
    Exit Function
End Function