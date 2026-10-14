Attribute VB_Name = "CSVUtils"
Option Explicit

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

Public Function ToLF(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    ToLF = s
End Function

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
    Dim i As Long
    Dim targetIndex As Long
    Dim batchIndex() As String
    Dim lineLen As Long
    Dim firstCommaPos As Long
    Dim tempContent As String
    Dim auditFilePath As String
    Dim tsTemp As Object
    Dim tsAudit As Object
    Dim auditTransID As String
    Dim auditLine As String
    
    UpdateFinishedGoodsBatchRecord = False
    
    If Trim$(batchNumber) = "" Then
        MsgBox "Error: Batch Number cannot be empty", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetDataFilePath(FILE_FINISHED_GOODS_BATCH_NUMBER_MASTER)
    tempFilePath = filePath & ".tmp"
    
    If IsFileLocked(filePath) Then
        MsgBox "FinishedGoodsBatchNumberMasterFile.csv is currently locked by another process. Please try again.", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    fileContent = ReadCSVFile(filePath)
    
    If UBound(fileContent) < 1 Then
        MsgBox "Error: FinishedGoodsBatchNumberMasterFile.csv is empty or could not be read", vbCritical, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
    End If
    
    headers = ParseCSVLine(fileContent(0))
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare
    
    For i = 0 To UBound(headers)
        headerMap(Trim$(Replace(headers(i), """", ""))) = i
    Next i
    
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
    
    values = ParseCSVLineOptimized(fileContent(targetIndex))
    
    If UBound(values) < 0 Then
        MsgBox "Error: Unable to parse record for Batch Number '" & batchNumber & "'", vbExclamation, "UpdateFinishedGoodsBatchRecord"
        GoTo CleanExit
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
    
    values(headerMap(FGBNM_HEADER_SCANNED)) = ""
    
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
    
    fso.DeleteFile filePath
    fso.MoveFile tempFilePath, filePath
    
    auditFilePath = GetDataFilePath(FILE_FINISHED_GOODS_BATCH_NUMBER_AUDIT)
    auditTransID = GetNextTransID(auditFilePath)
    auditLine = updatedLine & ",""" & auditTransID & """"
    
    Set tsAudit = fso.OpenTextFile(auditFilePath, 8, True)
    
    tsAudit.WriteLine auditLine
    tsAudit.Close
    
    UpdateFinishedGoodsBatchRecord = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "UpdateFinishedGoodsBatchRecord"
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(tempFilePath) Then fso.DeleteFile tempFilePath
    End If
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    If Not tsTemp Is Nothing Then Set tsTemp = Nothing
    If Not tsAudit Is Nothing Then Set tsAudit = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    If Not headerMap Is Nothing Then Set headerMap = Nothing
    Exit Function
End Function
