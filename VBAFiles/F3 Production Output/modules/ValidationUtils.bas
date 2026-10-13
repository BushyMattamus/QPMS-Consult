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

' ...existing code...

' ========================================================================
' Validate and handle batch status with user prompts for Goods Issue
' ========================================================================
Public Function ValidateAndHandleBatchStatusGoodsIssue(ByVal batchNumber As String, _
                                            ByVal callingProcedure As String, _
                                            ByRef requiresStatusUpdate As Boolean) As BatchStatusResult
    On Error GoTo ErrHandler
    
    Dim config As BatchStatusConfig
    Dim result As BatchStatusResult
    Dim currentStatus As String
    Dim response As VbMsgBoxResult
    Dim tempPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    config.targetFile = FILE_FINISHED_GOODS_BATCH_MASTER
    config.batchNumber = batchNumber
    config.allowStatuses = FGB_STATUS_PRINTED & "," & FGB_STATUS_INPROGRESS
    config.warnStatuses = FGB_STATUS_CREATED & "," & FGB_STATUS_COMPLETE
    config.stopStatuses = FGB_STATUS_REJECTED & "," & FGB_STATUS_CANCELLED
    
    tempPath = CreateTemporaryFileCopy(config.targetFile)
    
    If tempPath = "" Then
        ValidateAndHandleBatchStatusGoodsIssue = bsrError
        requiresStatusUpdate = False
        GoTo CleanExit
    End If
    
    currentStatus = QueryBatchStatus(tempPath, config.batchNumber)
    
    If currentStatus = "" Then
        MsgBox "Batch Number " & batchNumber & " not found.", vbExclamation, "Batch Not Found"
        ValidateAndHandleBatchStatusGoodsIssue = bsrNotFound
        requiresStatusUpdate = False
        GoTo CleanExit
    End If
    
    requiresStatusUpdate = (UCase$(Trim$(currentStatus)) <> UCase$(FGB_STATUS_INPROGRESS))
    
    result = EvaluateStatus(currentStatus, config.allowStatuses, config.warnStatuses, config.stopStatuses)
    
    Select Case result
        Case bsrAllow
            ValidateAndHandleBatchStatusGoodsIssue = bsrAllow
            
        Case bsrWarn
            response = MsgBox("Batch status is " & currentStatus & ". Click Yes to continue issuing goods, or No to cancel.", _
                            vbYesNo + vbExclamation, callingProcedure)
            If response = vbYes Then
                ValidateAndHandleBatchStatusGoodsIssue = bsrAllow
            Else
                ValidateAndHandleBatchStatusGoodsIssue = bsrWarn
                requiresStatusUpdate = False
            End If
            
        Case bsrStop
            MsgBox "Batch status is " & currentStatus & ". Cannot issue goods to this batch." & vbCrLf & _
                   "To change the status, use F1 Finished Goods Batch solution.", _
                   vbExclamation, "Cannot Issue Goods"
            ValidateAndHandleBatchStatusGoodsIssue = bsrStop
            requiresStatusUpdate = False
            
        Case bsrError
            MsgBox "Error validating batch status.", vbCritical, callingProcedure
            ValidateAndHandleBatchStatusGoodsIssue = bsrError
            requiresStatusUpdate = False
    End Select
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, callingProcedure
    ValidateAndHandleBatchStatusGoodsIssue = bsrError
    requiresStatusUpdate = False
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
' Validate and handle batch status with user prompts for Production Output
' ========================================================================
Public Function ValidateAndHandleBatchStatusProductionOutput(ByVal batchNumber As String, _
                                            ByVal callingProcedure As String, _
                                            ByRef requiresStatusUpdate As Boolean) As BatchStatusResult
    On Error GoTo ErrHandler
    
    Dim config As BatchStatusConfig
    Dim result As BatchStatusResult
    Dim currentStatus As String
    Dim response As VbMsgBoxResult
    Dim tempPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    config.targetFile = FILE_FINISHED_GOODS_BATCH_MASTER
    config.batchNumber = batchNumber
    config.allowStatuses = FGB_STATUS_INPROGRESS
    config.warnStatuses = FGB_STATUS_CREATED & "," & FGB_STATUS_PRINTED & "," & FGB_STATUS_REJECTED
    config.stopStatuses = FGB_STATUS_CANCELLED & "," & FGB_STATUS_COMPLETE
    
    tempPath = CreateTemporaryFileCopy(config.targetFile)
    
    If tempPath = "" Then
        ValidateAndHandleBatchStatusProductionOutput = bsrError
        requiresStatusUpdate = False
        GoTo CleanExit
    End If
    
    currentStatus = QueryBatchStatus(tempPath, config.batchNumber)
    
    If currentStatus = "" Then
        MsgBox "Batch Number " & batchNumber & " not found.", vbExclamation, "Batch Not Found"
        ValidateAndHandleBatchStatusProductionOutput = bsrNotFound
        requiresStatusUpdate = False
        GoTo CleanExit
    End If
    
    requiresStatusUpdate = (UCase$(Trim$(currentStatus)) <> UCase$(FGB_STATUS_COMPLETE))
    
    result = EvaluateStatus(currentStatus, config.allowStatuses, config.warnStatuses, config.stopStatuses)
    
    Select Case result
        Case bsrAllow
            ValidateAndHandleBatchStatusProductionOutput = bsrAllow
            
        Case bsrWarn
            response = MsgBox("Batch status is " & currentStatus & ". Click Yes to continue saving production output, or No to cancel.", _
                            vbYesNo + vbExclamation, callingProcedure)
            If response = vbYes Then
                ValidateAndHandleBatchStatusProductionOutput = bsrAllow
            Else
                ValidateAndHandleBatchStatusProductionOutput = bsrWarn
                requiresStatusUpdate = False
            End If
            
        Case bsrStop
            If UCase$(Trim$(currentStatus)) = UCase$(FGB_STATUS_COMPLETE) Then
                ValidateAndHandleBatchStatusProductionOutput = bsrAllow
                requiresStatusUpdate = False
            Else
                MsgBox "Batch status is " & currentStatus & ". Cannot save production output for this batch." & vbCrLf & _
                       "To change the status, use F1 Finished Goods Batch solution.", _
                       vbExclamation, "Cannot Save Production Output"
                ValidateAndHandleBatchStatusProductionOutput = bsrStop
                requiresStatusUpdate = False
            End If
            
        Case bsrError
            MsgBox "Error validating batch status.", vbCritical, callingProcedure
            ValidateAndHandleBatchStatusProductionOutput = bsrError
            requiresStatusUpdate = False
    End Select
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, callingProcedure
    ValidateAndHandleBatchStatusProductionOutput = bsrError
    requiresStatusUpdate = False
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
