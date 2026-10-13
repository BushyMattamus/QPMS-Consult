Attribute VB_Name = "Helpers"
Option Explicit

Public startTime As Double

' CSV parsing helper (handles quoted fields with commas)
Public Function ParseCSVLine(ByVal line As String) As String()
    Dim regex As Object
    Dim matches As Object
    Dim result() As String
    Dim i As Long
    Dim val As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(""([^""]|"""")*""|[^,]*)(,|$)"
    regex.Global = True
    
    Set matches = regex.Execute(line)
    ReDim result(0 To matches.Count - 1)
    
    For i = 0 To matches.Count - 1
        val = matches(i).SubMatches(0)
        
        ' Remove outer quotes if present ---
        If Left(val, 1) = """" And Right(val, 1) = """" Then
            val = Mid(val, 2, Len(val) - 2)
            val = Replace(val, """""", """")
        End If
        
        result(i) = val
    Next i
    
    ParseCSVLine = result
End Function

Function GetFieldByLabel(sheetName As String, labelCol As String, labelName As String) As Range
    Dim ws As Worksheet, foundCell As Range
    Set ws = ThisWorkbook.Sheets(sheetName)
    
      
    With ws.Columns(labelCol)
        Set foundCell = .Find(What:=labelName, LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    
    If foundCell Is Nothing Then
        MsgBox "Label not found on sheet '" & sheetName & "': " & labelName, vbCritical
        Stop
    Else
        Set GetFieldByLabel = foundCell.Offset(0, 1)
    End If
End Function

' Helper: Clears PMI number, PMI version, Shelf life fields, Manufacturing date
Public Sub ClearProductFields()

    Dim sheetName As String
    sheetName = ActiveSheet.Name
        
    On Error Resume Next ' In case a label is missing
    If sheetName = "Batch Number Report" Then
        Range("D5:D7").ClearContents
    ElseIf sheetName = "Batch Number Generation" Then
      Range("D8:D13").ClearContents
    ElseIf sheetName = "Batch Number Maintenance" Then
    Range("D7:D14").ClearContents
    
    End If
         
    On Error GoTo 0

End Sub

Public Sub ClearFields()

    Dim sheetName As String
    sheetName = ActiveSheet.Name
        
    On Error Resume Next ' In case a label is missing
    If sheetName = "Batch Number Report" Then
    Range("D5:D7").ClearContents
    
    ElseIf sheetName = "Batch Number Generation" Then
    Range("D7:D13").ClearContents
    
        
    ElseIf sheetName = "Batch Number Maintenance" Then
    Range("D4:D15").ClearContents
    End If
End Sub

Sub AppendLineToFile(filePath As String, lineText As String)
    On Error GoTo ErrHandler

    Dim fso As Object
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 8, True)  ' 8 = ForAppending
    ts.WriteLine lineText
    ts.Close

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "AppendLineToFile"
    GoTo CleanExit

CleanExit:
    If Not ts Is Nothing Then Set ts = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    Exit Sub
End Sub

Function GetNextTransID(auditCSVPath As String) As String
    Dim fileNum As Integer
    Dim lastLine As String
    Dim transId As String
    Dim numPart As Long
    
    ' --- If file doesn't exist, start from B000001 ---
    If Dir(auditCSVPath) = "" Then
        GetNextTransID = "B000001"
        Exit Function
    End If
    
    fileNum = FreeFile
    Open auditCSVPath For Input As #fileNum
    
    ' --- Read only the last line ---
    lastLine = ""
    Do While Not EOF(fileNum)
        Line Input #fileNum, lastLine
    Loop
    Close #fileNum
    
    ' --- If file is empty, start at B000001 ---
    If Trim(lastLine) = "" Then
        GetNextTransID = "B000001"
        Exit Function
    End If
    
    ' --- Split into columns ---
    Dim fields As Variant
    fields = Split(lastLine, ",")
    
    ' --- If the last column is missing, reset to B000001 ---
    If UBound(fields) < 15 Then
        GetNextTransID = "B000001"
        Exit Function
    End If
    
    ' --- Trans ID is last value ---
    transId = Trim(fields(UBound(fields)))
    
    ' --- Remove surrounding quotes if present ---
    If Left(transId, 1) = """" And Right(transId, 1) = """" Then
        transId = Mid(transId, 2, Len(transId) - 2)
    End If
    
    ' --- Unescape doubled quotes ---
    transId = Replace(transId, """""", """")
    
    ' --- Validate the pattern (B + 6 digits) ---
    If Len(transId) <> 7 Or Left$(transId, 1) <> "B" Or Not IsNumeric(Mid$(transId, 2)) Then
        GetNextTransID = "B000001"
        Exit Function
    End If
    
    ' --- Extract numeric part and increment ---
    numPart = CLng(Mid$(transId, 2)) + 1
    
    ' --- Build next Trans_ID ---
    GetNextTransID = "B" & Format$(numPart, "000000")
End Function

' QuickSort helper to sort parsed CSV rows by PMI (column 1)
Sub QuickSortParsed(ByRef data As Variant, ByRef idx As Variant, _
                            ByVal first As Long, ByVal last As Long)

    Dim low As Long, high As Long
    Dim pivot As String, temp As Long

    low = first
    high = last
    pivot = UCase$(data(idx((first + last) \ 2), 1))   ' PMI column

    Do While low <= high
        Do While UCase$(data(idx(low), 1)) < pivot
            low = low + 1
        Loop

        Do While UCase$(data(idx(high), 1)) > pivot
            high = high - 1
        Loop

        If low <= high Then
            temp = idx(low)
            idx(low) = idx(high)
            idx(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortParsed data, idx, first, high
    If low < last Then QuickSortParsed data, idx, low, last

End Sub

' Convert dd/mm/yyyy to YYYY-MM-DD (ANSI format)
Public Function DateToANSI(ByVal dateStr As String) As String
    On Error GoTo ErrorHandler
    
    Dim dateParts() As String
    dateParts = Split(Trim$(dateStr), "/")
    
    If UBound(dateParts) = 2 Then
        DateToANSI = dateParts(2) & "-" & _
                     Right$("0" & dateParts(1), 2) & "-" & _
                     Right$("0" & dateParts(0), 2)
    Else
        DateToANSI = ""
    End If
    Exit Function
    
ErrorHandler:
    DateToANSI = ""
End Function

' Convert YYYY-MM-DD (ANSI format) to dd/mm/yyyy
Public Function ANSIToDate(ByVal ansiDate As String) As String
    On Error GoTo ErrorHandler
    
    Dim dateParts() As String
    dateParts = Split(Trim$(ansiDate), "-")
    
    If UBound(dateParts) = 2 Then
        ANSIToDate = Right$("0" & dateParts(2), 2) & "/" & _
                     Right$("0" & dateParts(1), 2) & "/" & _
                     dateParts(0)
    Else
        ANSIToDate = ""
    End If
    Exit Function
    
ErrorHandler:
    ANSIToDate = ""
End Function

Public Function IsFileLocked(ByVal filePath As String) As Boolean
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open filePath For Binary Access Write Lock Read Write As #ff
    If Err.Number <> 0 Then
        IsFileLocked = True
        Err.Clear
    Else
        IsFileLocked = False
        Close #ff
    End If
    On Error GoTo 0
End Function

' Adds a timestamp to the ANSI formatted date
Public Function GetTransactionTimestamp() As String
    GetTransactionTimestamp = DateToANSI(Format(Now, "dd/mm/yyyy")) & " " & Format(Now, "hh:mm:ss")
End Function

' Timer functions for performance measurement
Public Sub StartTimer()
    startTime = Timer
End Sub

Public Sub EndTimer()
    Dim elapsedSeconds As Double
    Dim elapsedMinutes As Double
    Dim elapsedHours As Double
    
    elapsedSeconds = Timer - startTime
    elapsedMinutes = elapsedSeconds / 60
    elapsedHours = elapsedMinutes / 60
    
    Debug.Print "Elapsed time: " & Format(elapsedSeconds, "0.00") & " seconds"
    
    If elapsedMinutes >= 1 Then
        Debug.Print "Elapsed time: " & Format(elapsedMinutes, "0.00") & " minutes"
    End If
    
    If elapsedHours >= 1 Then
        Debug.Print "Elapsed time: " & Format(elapsedHours, "0.00") & " hours"
    End If
End Sub
