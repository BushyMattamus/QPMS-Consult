Attribute VB_Name = "Functions"
Option Explicit

Public startTime As Double


' --- Helpers ---

Public Function FindHeaderCol(ws As Worksheet, headerRow As Long, headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(CStr(ws.Cells(headerRow, c).value), headerText, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Public Function RM_Pad4(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        RM_Pad4 = "0000"
    Else
        ' keep digits padded; if not purely digits, just return the text as-is
        If t Like "*[!0-9]*" Then
            RM_Pad4 = t
        Else
            RM_Pad4 = Right$("0000" & t, 4)
        End If
    End If
End Function

Public Function MakeTempSheetName(baseName As String) As String
    Dim nm As String, i As Long
    nm = baseName
    i = 1
    Do While SheetExists(nm)
        nm = baseName & "_" & i
        i = i + 1
    Loop
    MakeTempSheetName = nm
End Function

Public Function SheetExists(shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shtName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

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

' Adds a timestamp to the ANSI formatted date
Public Function GetTransactionTimestamp() As String
    GetTransactionTimestamp = DateToANSI(Format(Now, "dd/mm/yyyy")) & " " & Format(Now, "hh:mm:ss")
End Function

' Wrap the output dates in the following to ensure consistent date format display
' example: SetDateStringInCell Range("D11"), ANSIToDate(cleanValue)
Public Sub SetDateStringInCell(Target As Range, dateStr As String)
    With Target
        .NumberFormat = "@"
        .value = dateStr
        .HorizontalAlignment = xlLeft
    End With
    Target.NumberFormat = "dd/mm/yyyy"
End Sub

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

Public Sub DebugListCSVHeaders()
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim headerLine As String
    Dim headers() As String
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not fso.FileExists(filePath) Then
        Debug.Print "File not found: " & filePath
        GoTo CleanExit
    End If
    
    Set ts = fso.OpenTextFile(filePath, 1)
    headerLine = ts.ReadLine
    ts.Close
    
    headers = Split(headerLine, ",")
    
    Debug.Print "=== QC_Number.csv Headers (0-based index) ==="
    Debug.Print "Total columns: " & (UBound(headers) + 1)
    Debug.Print ""
    
    For i = 0 To UBound(headers)
        Debug.Print "Index " & i & ": " & Replace(headers(i), """", "")
    Next i
    
    Debug.Print ""
    Debug.Print "=== Looking for specific headers ==="
    Debug.Print "Item_Cost index: " & FindHeaderIndex(headers, "Item_Cost")
    Debug.Print "BUoM_Cost index: " & FindHeaderIndex(headers, "BUoM_Cost")
    Debug.Print "PO_Number index: " & FindHeaderIndex(headers, "PO_Number")
    
CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

Private Function FindHeaderIndex(headers() As String, headerName As String) As Long
    Dim i As Long
    For i = 0 To UBound(headers)
        If Trim$(Replace(headers(i), """", "")) = headerName Then
            FindHeaderIndex = i
            Exit Function
        End If
    Next i
    FindHeaderIndex = -1
End Function

Public Sub DebugTestHeaderIndices()
    Debug.Print "=== Testing GetHeaderIndex0 ==="
    Debug.Print "idxItemCost = " & GetHeaderIndex0(HEADER_ITEM_COST) & " (expecting 42)"
    Debug.Print "idxBuomCost = " & GetHeaderIndex0(HEADER_BUOM_COST) & " (expecting 43)"
    Debug.Print "idxPONumber = " & GetHeaderIndex0(HEADER_PO_NUMBER) & " (expecting 44)"
    Debug.Print "idxTotalQty = " & GetHeaderIndex0(HEADER_TOTAL_QTY) & " (expecting 11)"
    Debug.Print "idxQCBalance = " & GetHeaderIndex0(HEADER_QC_BALANCE) & " (expecting 20)"
    Debug.Print ""
    Debug.Print "HEADER_ITEM_COST constant = '" & HEADER_ITEM_COST & "'"
    Debug.Print "HEADER_BUOM_COST constant = '" & HEADER_BUOM_COST & "'"
    Debug.Print "HEADER_PO_NUMBER constant = '" & HEADER_PO_NUMBER & "'"
End Sub