Attribute VB_Name = "BatchNumberGenerator"
' BatchNumberGenerator.bas

Option Explicit

' ========================================================================
' BATCH NUMBER GENERATION
' ========================================================================
' Generates batch numbers by querying the master CSV file to find the
' highest sequence number for today's date, then incrementing it.
' Format: [MonthLetter][Day][Sequence][Year] (e.g., A2201026)
' ========================================================================

Public Function GenerateBatchNumber() As String
    On Error GoTo ErrHandler

    Dim today As Date
    Dim monthLetter As String
    Dim dayStr As String
    Dim yearStr As String
    Dim seqNum As Long
    Dim seqStr As String

    today = Date

    monthLetter = GetMonthLetter(today)
    dayStr = Format$(Day(today), "00")
    yearStr = Format$(Year(today) Mod 100, "00")

    seqNum = GetNextSequenceNumber(today)
    seqStr = Format$(seqNum, "00")

    GenerateBatchNumber = monthLetter & dayStr & seqStr & yearStr
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GenerateBatchNumber"
    GenerateBatchNumber = vbNullString
    GoTo CleanExit

CleanExit:
    Exit Function
End Function

' ========================================================================
' Returns the month letter using the mapping table
' ========================================================================
Private Function GetMonthLetter(ByVal d As Date) As String
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim m As Long
    Dim result As Variant

    Set ws = ThisWorkbook.Sheets("Batch number")
    m = Month(d)

    result = Application.VLookup(m, ws.Range("A:B"), 2, False)

    If IsError(result) Then
        Err.Raise vbObjectError + 1001, , "Month mapping not found for month " & m
    End If

    GetMonthLetter = CStr(result)
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetMonthLetter"
    GetMonthLetter = ""
    GoTo CleanExit

CleanExit:
    Set ws = Nothing
    Exit Function
End Function

' ========================================================================
' Returns the next sequence number (1-99) for the given date by querying
' the master CSV file for all batch numbers created on that date
' ========================================================================
Private Function GetNextSequenceNumber(ByVal d As Date) As Long
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim rs As Object
    Dim tempFilePath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim batchNum As String
    Dim seqStr As String
    Dim seqNum As Long
    Dim maxSeq As Long
    Dim targetDate As String
    Dim fso As Object

    maxSeq = 0

    ' --- Create temporary copy for safe reading ---
    tempFilePath = CreateTemporaryFileCopy(FILE_FINISHED_GOODS_BATCH_NUMBER_MASTER)
    If tempFilePath = "" Then
        MsgBox "Unable to create temporary file for sequence number lookup.", vbCritical, "GetNextSequenceNumber"
        GetNextSequenceNumber = 1
        GoTo CleanExit
    End If

    ' --- Extract folder and filename ---
    folderPath = Left(tempFilePath, InStrRev(tempFilePath, "\") - 1)
    fileName = Mid(tempFilePath, InStrRev(tempFilePath, "\") + 1)

    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(folderPath)
    If conn Is Nothing Then
        GetNextSequenceNumber = 1
        GoTo CleanExit
    End If

    ' --- Query for batch numbers created on target date ---
    Set rs = CreateObject("ADODB.Recordset")
    targetDate = Format(d, "yyyy-mm-dd")
    
    sql = "SELECT " & FGBNM_HEADER_BATCH_NUMBER & ", " & FGBNM_HEADER_DATE_CREATED & _
          " FROM [" & fileName & "] " & _
          " WHERE " & FGBNM_HEADER_DATE_CREATED & " = '" & targetDate & "'"

    rs.Open sql, conn

    ' --- Loop through results and find max sequence ---
    Do While Not rs.EOF
        batchNum = CStr(rs.fields(FGBNM_HEADER_BATCH_NUMBER).value)
        
        If Len(batchNum) >= 6 Then
            seqStr = Mid$(batchNum, 4, 2)
            
            If IsNumeric(seqStr) Then
                seqNum = CLng(seqStr)
                If seqNum > maxSeq Then maxSeq = seqNum
            End If
        End If
        
        rs.MoveNext
    Loop

    GetNextSequenceNumber = maxSeq + 1
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetNextSequenceNumber"
    GetNextSequenceNumber = 1
    GoTo CleanExit

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If

    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If

    ' --- Clean up temporary file ---
    If tempFilePath <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(tempFilePath) Then
            On Error Resume Next
            fso.DeleteFile tempFilePath, True
            On Error GoTo 0
        End If
        Set fso = Nothing
    End If

    Exit Function
End Function
