Attribute VB_Name = "ValidationUtils"
' ========================================================================
' VALIDATION FUNCTIONS
' ========================================================================

Option Explicit

Public Enum BatchStatusResult
    bsrAllow = 0
    bsrWarn = 1
    bsrStop = 2
    bsrNotFound = 3
    bsrError = 4
End Enum

Public Type BatchStatusConfig
    targetFile As String
    batchNumber As String
    allowStatuses As String
    warnStatuses As String
    stopStatuses As String
End Type

' ========================================================================
' Main validation function
' ========================================================================
Public Function ValidateBatchStatus(config As BatchStatusConfig) As BatchStatusResult
    On Error GoTo ErrHandler
    
    Dim tempPath As String
    Dim currentStatus As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    tempPath = CreateTemporaryFileCopy(config.targetFile)
    
    If tempPath = "" Then
        ValidateBatchStatus = bsrError
        GoTo CleanExit
    End If
    
    currentStatus = QueryBatchStatus(tempPath, config.batchNumber)
    
    If currentStatus = "" Then
        ValidateBatchStatus = bsrNotFound
        GoTo CleanExit
    End If
    
    ValidateBatchStatus = EvaluateStatus(currentStatus, config.allowStatuses, config.warnStatuses, config.stopStatuses)
    
    GoTo CleanExit

ErrHandler:
    ValidateBatchStatus = bsrError
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' Query batch status from CSV using ADO
' ========================================================================
Private Function QueryBatchStatus(ByVal filePath As String, ByVal batchNumber As String) As String
    On Error GoTo ErrHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rs = CreateObject("ADODB.Recordset")
    
    folderPath = fso.GetParentFolderName(filePath)
    fileName = fso.GetFileName(filePath)
    
    Set conn = OpenADOConnection(folderPath)
    
    If conn Is Nothing Then
        QueryBatchStatus = ""
        GoTo CleanExit
    End If
    
    sql = "SELECT [" & FGBNM_HEADER_BATCH_STATUS & "] " & _
          "FROM [" & fileName & "] " & _
          "WHERE [" & FGBNM_HEADER_BATCH_NUMBER & "] = '" & Replace(batchNumber, "'", "''") & "'"
    
    rs.Open sql, conn
    
    If Not rs.EOF Then
        QueryBatchStatus = Trim$(rs.Fields(0).value & "")
    Else
        QueryBatchStatus = ""
    End If
    
    GoTo CleanExit

ErrHandler:
    QueryBatchStatus = ""
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set rs = Nothing
    Set conn = Nothing
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' Evaluate status against allowed/warn/stop lists
' ========================================================================
Private Function EvaluateStatus(ByVal currentStatus As String, _
                                ByVal allowStatuses As String, _
                                ByVal warnStatuses As String, _
                                ByVal stopStatuses As String) As BatchStatusResult
    
    currentStatus = Trim$(UCase$(currentStatus))
    
    If Len(allowStatuses) > 0 Then
        If IsStatusInList(currentStatus, allowStatuses) Then
            EvaluateStatus = bsrAllow
            Exit Function
        End If
    End If
    
    If Len(warnStatuses) > 0 Then
        If IsStatusInList(currentStatus, warnStatuses) Then
            EvaluateStatus = bsrWarn
            Exit Function
        End If
    End If
    
    If Len(stopStatuses) > 0 Then
        If IsStatusInList(currentStatus, stopStatuses) Then
            EvaluateStatus = bsrStop
            Exit Function
        End If
    End If
    
    EvaluateStatus = bsrWarn
End Function

' ========================================================================
' Check if a status exists in a comma-separated list
' ========================================================================
Private Function IsStatusInList(ByVal status As String, ByVal statusList As String) As Boolean
    Dim statuses() As String
    Dim i As Long
    
    status = Trim$(UCase$(status))
    statusList = Trim$(UCase$(statusList))
    
    statuses = Split(statusList, ",")
    
    For i = LBound(statuses) To UBound(statuses)
        If StrComp(status, Trim$(statuses(i)), vbTextCompare) = 0 Then
            IsStatusInList = True
            Exit Function
        End If
    Next i
    
    IsStatusInList = False
End Function

' ========================================================================
' Validate and handle batch status with user prompts for New Prints
' ========================================================================
Public Function ValidateAndHandleBatchStatusNew(ByVal batchNumber As String, _
                                            ByVal callingProcedure As String) As BatchStatusResult
    On Error GoTo ErrHandler
    
    Dim config As BatchStatusConfig
    Dim result As BatchStatusResult
    Dim response As VbMsgBoxResult
    
    config.targetFile = FILE_FINISHED_GOODS_BATCH_MASTER
    config.batchNumber = batchNumber
    config.allowStatuses = FGB_STATUS_CREATED & "," & FGB_STATUS_PRINTED
    config.warnStatuses = FGB_STATUS_INPROGRESS & "," & FGB_STATUS_COMPLETE
    config.stopStatuses = FGB_STATUS_REJECTED & "," & FGB_STATUS_CANCELLED
    
    result = ValidateBatchStatus(config)
    
    Select Case result
        Case bsrAllow
            ValidateAndHandleBatchStatusNew = bsrAllow
            
        Case bsrWarn
            response = MsgBox("Batch status is In Progress or Completed. Click Yes to continue loading, or No to cancel.", _
                            vbYesNo + vbExclamation, callingProcedure)
            If response = vbYes Then
                ValidateAndHandleBatchStatusNew = bsrAllow
            Else
                ValidateAndHandleBatchStatusNew = bsrWarn
            End If
            
        Case bsrStop
            MsgBox "Batch status is Rejected or Cancelled. To change the status, use F1 Finished Goods Batch solution.", _
                   vbExclamation, "Cannot Load Batch"
            ValidateAndHandleBatchStatusNew = bsrStop
            
        Case bsrNotFound
            MsgBox "Batch Number " & batchNumber & " not found.", vbExclamation, "Batch Not Found"
            ValidateAndHandleBatchStatusNew = bsrNotFound
            
        Case bsrError
            MsgBox "Error validating batch status.", vbCritical, callingProcedure
            ValidateAndHandleBatchStatusNew = bsrError
    End Select
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, callingProcedure
    ValidateAndHandleBatchStatusNew = bsrError
    GoTo CleanExit

CleanExit:
    Exit Function
End Function

' ========================================================================
' Validate and handle batch status with user prompts for Rerints
' ========================================================================
Public Function ValidateAndHandleBatchStatusReprint(ByVal batchNumber As String, _
                                            ByVal callingProcedure As String) As BatchStatusResult
    On Error GoTo ErrHandler
    
    Dim config As BatchStatusConfig
    Dim result As BatchStatusResult
    Dim response As VbMsgBoxResult
    
    config.targetFile = FILE_FINISHED_GOODS_BATCH_MASTER
    config.batchNumber = batchNumber
    config.allowStatuses = FGB_STATUS_CREATED & "," & FGB_STATUS_PRINTED
    config.warnStatuses = FGB_STATUS_INPROGRESS & "," & FGB_STATUS_COMPLETE & "," & FGB_STATUS_REJECTED
    config.stopStatuses = FGB_STATUS_CANCELLED
    
    result = ValidateBatchStatus(config)
    
    Select Case result
        Case bsrAllow
            ValidateAndHandleBatchStatusReprint = bsrAllow
            
        Case bsrWarn
            response = MsgBox("Batch status is In Progress, Completed or Rejected. Click Yes to continue loading, or No to cancel.", _
                            vbYesNo + vbExclamation, callingProcedure)
            If response = vbYes Then
                ValidateAndHandleBatchStatusReprint = bsrAllow
            Else
                ValidateAndHandleBatchStatusReprint = bsrWarn
            End If
            
        Case bsrStop
            MsgBox "Batch status is Cancelled. To change the status, use F1 Finished Goods Batch solution.", _
                   vbExclamation, "Cannot Load Batch"
            ValidateAndHandleBatchStatusReprint = bsrStop
            
        Case bsrNotFound
            MsgBox "Batch Number " & batchNumber & " not found.", vbExclamation, "Batch Not Found"
            ValidateAndHandleBatchStatusReprint = bsrNotFound
            
        Case bsrError
            MsgBox "Error validating batch status.", vbCritical, callingProcedure
            ValidateAndHandleBatchStatusReprint = bsrError
    End Select
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, callingProcedure
    ValidateAndHandleBatchStatusReprint = bsrError
    GoTo CleanExit

CleanExit:
    Exit Function
End Function

' ========================================================================
' Validates that exactly one template file exists for the given pattern
' ========================================================================
Public Function ValidateUniqueTemplateFile(ByVal folderPath As String, _
                                           ByVal pmiNumber As String, _
                                           ByVal filePattern As String, _
                                           ByVal templateType As String, _
                                           ByRef fileName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim fileCount As Integer
    Dim pmiPrefixCount As Integer
    Dim nextFileName As String
    Dim currentFileName As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    ValidateUniqueTemplateFile = False
    fileName = ""
    fileCount = 0
    pmiPrefixCount = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        GoTo CleanExit
    End If

    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        currentFileName = file.Name
        If Left(currentFileName, Len(pmiNumber)) = pmiNumber Then
            pmiPrefixCount = pmiPrefixCount + 1
            
            If currentFileName Like filePattern Then
                fileCount = fileCount + 1
                If fileName = "" Then
                    fileName = currentFileName
                End If
            End If
        End If
    Next file
    
    If pmiPrefixCount = 0 Then
        MsgBox templateType & " template not found for PMI Number: " & pmiNumber, vbExclamation, "Validation Error"
        GoTo CleanExit
    End If
    
    If pmiPrefixCount > 1 Then
        MsgBox "Multiple files found starting with PMI Number: " & pmiNumber & vbCrLf & _
               "Found " & pmiPrefixCount & " file(s) in the " & templateType & " folder." & vbCrLf & _
               "Please ensure only one file exists for this PMI Number.", vbExclamation, "Validation Error"
        GoTo CleanExit
    End If
    
    If fileCount > 1 Then
        MsgBox "Multiple " & templateType & " template files found matching pattern: " & filePattern & vbCrLf & _
               "Expected exactly one file for PMI Number: " & pmiNumber & vbCrLf & _
               "Please remove duplicate " & templateType & " templates from the folder.", vbExclamation, "Validation Error"
        GoTo CleanExit
    End If
    
    If fileCount = 0 Then
        MsgBox templateType & " template not found matching pattern: " & filePattern & vbCrLf & _
               "For PMI Number: " & pmiNumber, vbExclamation, "Validation Error"
        GoTo CleanExit
    End If
    
    ValidateUniqueTemplateFile = True
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ValidateUniqueTemplateFile"
    ValidateUniqueTemplateFile = False
    GoTo CleanExit

CleanExit:
    Exit Function
End Function

' ========================================================================
' Validates that the PMI version in the label document matches expected version
' ========================================================================
Public Function ValidateLabelPMIVersion(ByVal wdDoc As Object, _
                                        ByVal wdApp As Object, _
                                        ByVal pmiNumber As String, _
                                        ByVal expectedVersion As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim docContent As Object
    Dim docText As String
    Dim pmiPattern As String
    Dim extractedVersion As String
    Dim ch As String
    Dim startPos As Long
    Dim dotPos As Long
    Dim versionEndPos As Long
    Dim fullPmiText As String
    Dim i As Long
    
    ValidateLabelPMIVersion = False
    
    Set docContent = wdDoc.Content
    docText = docContent.Text
    
    pmiPattern = pmiNumber & "."
    startPos = InStr(1, docText, pmiPattern, vbTextCompare)
    
    If startPos = 0 Then
        wdDoc.Close SaveChanges:=False
        wdApp.Quit SaveChanges:=False
        MsgBox "Error: PMI Number '" & pmiNumber & "' not found in label document.", vbExclamation, "ValidateLabelPMIVersion"
        GoTo CleanExit
    End If
    
    dotPos = startPos + Len(pmiNumber)
    versionEndPos = dotPos + 1
    
    For i = dotPos + 1 To Len(docText)
        ch = Mid$(docText, i, 1)
        If IsNumeric(ch) Then
            versionEndPos = i
        Else
            Exit For
        End If
    Next i
    
    fullPmiText = Mid$(docText, startPos, versionEndPos - startPos + 1)
    extractedVersion = Mid$(fullPmiText, Len(pmiNumber) + 2)
    extractedVersion = Trim$(extractedVersion)
    
    If extractedVersion <> expectedVersion Then
        wdDoc.Close SaveChanges:=False
        wdApp.Quit SaveChanges:=False
        MsgBox "Version mismatch detected!" & vbCrLf & vbCrLf & _
               "Expected Version: " & expectedVersion & vbCrLf & _
               "Document Version: " & extractedVersion & vbCrLf & vbCrLf & _
               "Please verify the correct label document is being used.", vbExclamation, "ValidateLabelPMIVersion"
        GoTo CleanExit
    End If
    
    ValidateLabelPMIVersion = True
    GoTo CleanExit

ErrHandler:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit SaveChanges:=False
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ValidateLabelPMIVersion"
    ValidateLabelPMIVersion = False
    GoTo CleanExit

CleanExit:
    Set docContent = Nothing
    Exit Function
End Function

' ========================================================================
' Validates that manufacturing lot numbers are provided when label template requires them
' ========================================================================
Public Function ValidateLabelMfgLotNumbers(ByVal wdDoc As Object, _
                                           ByVal wdApp As Object, _
                                           ByVal mLotValue As String, _
                                           ByVal mLot2Value As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim docContent As Object
    Dim docText As String
    Dim hasMLot As Boolean
    Dim hasMLot2 As Boolean
    Dim missingValues As String
    
    ValidateLabelMfgLotNumbers = False
    
    Set docContent = wdDoc.Content
    docText = docContent.Text
    
    hasMLot = (InStr(1, docText, "?MLOT", vbTextCompare) > 0)
    hasMLot2 = (InStr(1, docText, "?MLOT2", vbTextCompare) > 0)
    
    missingValues = ""
    
    If hasMLot And Trim$(mLotValue) = "" Then
        missingValues = "Mfg Lot No"
    End If
    
    If hasMLot2 And Trim$(mLot2Value) = "" Then
        If missingValues <> "" Then
            missingValues = missingValues & " and "
        End If
        missingValues = missingValues & "Mfg Lot No 2"
    End If
    
    If missingValues <> "" Then
        wdDoc.Close SaveChanges:=False
        wdApp.Quit SaveChanges:=False
        MsgBox "The label template requires the following field(s) to be completed:" & vbCrLf & vbCrLf & _
               missingValues & vbCrLf & vbCrLf & _
               "Please enter the required manufacturing lot number(s) and try again.", _
               vbExclamation, "Validation Error"
        GoTo CleanExit
    End If
    
    ValidateLabelMfgLotNumbers = True
    GoTo CleanExit

ErrHandler:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit SaveChanges:=False
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ValidateLabelMfgLotNumbers"
    ValidateLabelMfgLotNumbers = False
    GoTo CleanExit

CleanExit:
    Set docContent = Nothing
    Exit Function
End Function

' ========================================================================
' Validates that the PMI version in the PMI document header matches expected version
' ========================================================================
Public Function ValidatePMIDocumentVersion(ByVal wdDoc As Object, _
                                          ByVal wdApp As Object, _
                                          ByVal expectedVersion As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim headerTable As Object
    Dim versionCell As Object
    Dim versionText As String
    Dim extractedVersion As String
    Dim cleanVersion As String
    Dim versionPos As Long
    Dim quickPos As Long
    Dim i As Long
    
    ValidatePMIDocumentVersion = False
    
    On Error Resume Next
    Set headerTable = wdDoc.Sections(1).headers(1).Range.Tables(1)
    On Error GoTo ErrHandler
    
    If headerTable Is Nothing Then
        wdDoc.Close SaveChanges:=False
        wdApp.Quit SaveChanges:=False
        MsgBox "Error: Unable to locate header table in PMI document.", vbExclamation, "ValidatePMIDocumentVersion"
        GoTo CleanExit
    End If
    
    Set versionCell = headerTable.Cell(1, 2).Range
    versionText = versionCell.Text
    versionText = Replace(versionText, vbCr, "")
    versionText = Replace(versionText, vbLf, "")
    versionText = Replace(versionText, Chr(13), "")
    versionText = Replace(versionText, Chr(7), "")
    versionText = Trim$(versionText)
    
    versionPos = InStr(1, versionText, "Version:", vbTextCompare)
    
    If versionPos > 0 Then
        extractedVersion = Mid$(versionText, versionPos + 8)
        extractedVersion = Trim$(extractedVersion)
        
        quickPos = InStr(1, extractedVersion, "Quick", vbTextCompare)
        If quickPos > 0 Then
            extractedVersion = Left$(extractedVersion, quickPos - 1)
        End If
        
        cleanVersion = ""
        For i = 1 To Len(extractedVersion)
            If Asc(Mid$(extractedVersion, i, 1)) >= 32 Then
                cleanVersion = cleanVersion & Mid$(extractedVersion, i, 1)
            End If
        Next i
        extractedVersion = Trim$(cleanVersion)
    Else
        wdDoc.Close SaveChanges:=False
        wdApp.Quit SaveChanges:=False
        MsgBox "Error: Version information not found in document header.", vbExclamation, "ValidatePMIDocumentVersion"
        GoTo CleanExit
    End If
    
    If extractedVersion <> expectedVersion Then
        wdDoc.Close SaveChanges:=False
        wdApp.Quit SaveChanges:=False
        MsgBox "Version mismatch detected!" & vbCrLf & vbCrLf & _
               "Expected Version: " & expectedVersion & vbCrLf & _
               "Document Version: " & extractedVersion & vbCrLf & vbCrLf & _
               "Please verify the correct PMI document is being used.", vbExclamation, "ValidatePMIDocumentVersion"
        GoTo CleanExit
    End If
    
    ValidatePMIDocumentVersion = True
    GoTo CleanExit

ErrHandler:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit SaveChanges:=False
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ValidatePMIDocumentVersion"
    ValidatePMIDocumentVersion = False
    GoTo CleanExit

CleanExit:
    Set versionCell = Nothing
    Set headerTable = Nothing
    Exit Function
End Function
