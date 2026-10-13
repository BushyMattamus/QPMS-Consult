Attribute VB_Name = "Functions"
Option Explicit

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

' Safely release object references (late-bound or early-bound)
Public Sub SafeRelease(ByRef obj As Object)
    On Error Resume Next
    If Not obj Is Nothing Then Set obj = Nothing
End Sub

' Centralized error message (extend later to log to file)
Public Sub StandardErrorReport(ByVal procName As String, ByVal errNum As Long, ByVal errDesc As String)
    MsgBox procName & " failed (" & errNum & "): " & errDesc, vbExclamation
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

' Wrap the output dates in the following to ensure consistent date format display
' example: SetDateStringInCell Range("D11"), ANSIToDate(cleanValue)
Public Sub SetDateStringInCell(target As Range, dateStr As String)
    With target
        .NumberFormat = "@"
        .value = dateStr
        .HorizontalAlignment = xlLeft
    End With
    target.NumberFormat = "dd/mm/yyyy"
End Sub

' Adds a timestamp to the ANSI formatted date
Public Function GetTransactionTimestamp() As String
    GetTransactionTimestamp = DateToANSI(Format(Now, "dd/mm/yyyy")) & " " & Format(Now, "hh:mm:ss")
End Function
