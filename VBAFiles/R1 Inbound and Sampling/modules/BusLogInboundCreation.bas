Attribute VB_Name = "BusLogInboundCreation"
Option Explicit

'Public Sub GetNextQCNumber()
'' **** NOTE: if this doesn't work, change the BOM_UTF8 = "" otherwise try ChrW(&HFEFF)
'    Dim conn As Object
'    Dim rs As Object
'    Dim filePath As String
'    Dim fileName As String
'    Dim sql As String
'    Dim lastQC As String
'    Dim currentYear As String
'    Dim lastYear As String
'    Dim seqNum As Long
'    Dim newQCValue As String
'    Dim BOM_UTF8 As String
'
'    On Error GoTo ErrorHandler
'
'    ' --- Initialize defaults ---
'    currentYear = Right$(CStr(Year(Date)), 2)
'    newQCValue = currentYear & "0001"
'    BOM_UTF8 = ChrW(&HFEFF) ' UTF-8 BOM character
'
'    ' --- Set file path ---
'    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
'
'    ' --- Check if file exists ---
'    If Dir(filePath) = "" Then
'        GoTo OutputResult
'    End If
'
'    ' --- Open ADO connection ---
'    Set conn = OpenADOConnection(Left(filePath, InStrRev(filePath, "\") - 1))
'    If conn Is Nothing Then
'        MsgBox "Failed to open ADO connection.", vbCritical, "GetNextQCNumber"
'        GoTo OutputResult
'    End If
'
'    ' --- SQL query to get the last QC_Number ---
'    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
'    sql = "SELECT TOP 1 [" & BOM_UTF8 & HEADER_QC_NUMBER & "] FROM [" & fileName & "] " & _
'          "WHERE [" & BOM_UTF8 & HEADER_QC_NUMBER & "] NOT LIKE '%[A-Za-z]%' " & _
'          "ORDER BY [" & BOM_UTF8 & HEADER_QC_NUMBER & "] DESC"
'
'    ' --- Execute query ---
'    Set rs = conn.Execute(sql)
'
'    ' --- Get the last QC_Number and calculate next ---
'    If Not rs.EOF Then
'        lastQC = Trim$(Replace(CStr(rs.Fields(HEADER_QC_NUMBER).value), """", ""))
'
'        If Len(lastQC) = 6 And IsNumeric(lastQC) Then
'            lastYear = Left$(lastQC, 2)
'
'            If lastYear = currentYear Then
'                seqNum = CLng(Right$(lastQC, 4)) + 1
'                newQCValue = currentYear & Format$(seqNum, "0000")
'            End If
'        End If
'    End If
'
'OutputResult:
'    ' --- Output to worksheet ---
'    With ThisWorkbook.Sheets(1).Range("D3")
'        .NumberFormat = "@"
'        .value = newQCValue
'    End With
'
'    GoTo CleanExit
'
'ErrorHandler:
'    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetNextQCNumber"
'
'CleanExit:
'    If Not rs Is Nothing Then
'        If Not rs.State = 0 Then rs.Close
'        Set rs = Nothing
'    End If
'    If Not conn Is Nothing Then
'        If Not conn.State = 0 Then conn.Close
'        Set conn = Nothing
'    End If
'    Debug.Print newQCValue
'End Sub

Public Sub GetNextQCNumber()
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim content As String
    Dim lines As Variant
    Dim i As Long
    Dim parts As Variant
    Dim lastQC As String
    Dim currentYear As String
    Dim lastYear As String
    Dim seqNum As Long
    Dim newQCValue As String

    On Error GoTo ErrorHandler

    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set ts = fso.OpenTextFile(filePath, 1)
    content = ts.ReadAll
    ts.Close

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines = Split(content, vbLf)

    currentYear = Right$(CStr(Year(Date)), 2)
    newQCValue = currentYear & "0001"

    For i = UBound(lines) To 1 Step -1
        If Trim$(lines(i)) <> "" Then
            parts = Split(lines(i), ",")
            If UBound(parts) >= 0 Then
                lastQC = Trim$(parts(0))
                lastQC = Replace(lastQC, """", "")
                
                If Len(lastQC) = 6 And IsNumeric(lastQC) Then
                    lastYear = Left$(lastQC, 2)
                    
                    If lastYear = currentYear Then
                        seqNum = CLng(Right$(lastQC, 4)) + 1
                        newQCValue = currentYear & Format$(seqNum, "0000")
                    End If
                    Exit For
                End If
            End If
        End If
    Next i

    With ThisWorkbook.Sheets(1).Range("D3")
        .NumberFormat = "@"
        .value = newQCValue
    End With

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in GetNextQCNumber: " & Err.Description, vbCritical, "Error"

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Sub
End Sub
Private Function KeepDigits(ByVal s As String) As String
    Dim i As Long, ch As String, outS As String
    outS = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then outS = outS & ch
    Next i
    KeepDigits = outS
End Function


Public Sub TestQCNumberConnection()
    Dim conn As Object
    Dim rs As Object
    Dim filePath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String

    On Error GoTo ErrorHandler

    ' --- Set file path ---
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)

    ' --- Check if file exists ---
    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbCritical, "TestQCNumberConnection"
        GoTo CleanExit
    End If

    Debug.Print "File Path: " & filePath

    ' --- Extract folder path ---
    folderPath = Left(filePath, InStrRev(filePath, "\") - 1)
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    Debug.Print "Folder Path: " & folderPath
    Debug.Print "File Name: " & fileName

    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(folderPath)
    If conn Is Nothing Then
        MsgBox "Failed to open ADO connection.", vbCritical, "TestQCNumberConnection"
        GoTo CleanExit
    End If

    Debug.Print "ADO Connection: SUCCESS"

    ' --- Simple test query to get all records ---
    sql = "SELECT * FROM [" & fileName & "]"
    Debug.Print "Query: " & sql

    ' --- Execute query ---
    Set rs = conn.Execute(sql)

    ' --- Display field names ---
    Debug.Print "Fields: "
    Dim fld As Object
    For Each fld In rs.Fields
        Debug.Print "  - " & fld.Name
    Next fld

    ' --- Display first 5 rows ---
    Debug.Print "First 5 rows:"
    Dim rowCount As Long
    rowCount = 0
    If Not rs.EOF Then
        While Not rs.EOF And rowCount < 5
            Debug.Print "  Row " & rowCount + 1 & ": " & rs.Fields(0).value
            rs.MoveNext
            rowCount = rowCount + 1
        Wend
    Else
        Debug.Print "  (No records found)"
    End If

    GoTo CleanExit

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "TestQCNumberConnection"

CleanExit:
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If Not conn.State = 0 Then conn.Close
        Set conn = Nothing
    End If
    Exit Sub
End Sub
