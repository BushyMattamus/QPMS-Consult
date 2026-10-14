Attribute VB_Name = "Functions"
Public startTime As Double

'Function GetGeneralPassword() As String
'    Dim ws As Worksheet
'    On Error Resume Next
'    Set ws = ThisWorkbook.Sheets("Admin")
'
'    ' Check if the sheet exists and C2 has a value
'    If Not ws Is Nothing Then
'        GetGeneralPassword = Trim(ws.Range("C2").value) ' Read password from C2
'    Else
'        GetGeneralPassword = "" ' Return empty if sheet not found
'    End If
'
'    On Error GoTo 0
'End Function
'
'Function GetLabelMaintenancePassword() As String
'    Dim ws As Worksheet
'    On Error Resume Next
'    Set ws = ThisWorkbook.Sheets("Admin")
'
'    ' Check if the sheet exists and C3 has a value
'    If Not ws Is Nothing Then
'        GetLabelMaintenancePassword = Trim(ws.Range("C3").value) ' Read password from C3
'    Else
'        GetLabelMaintenancePassword = "" ' Return empty if sheet not found
'    End If
'
'    On Error GoTo 0
'End Function
'
'Function GetAdminPassword() As String
'    Dim ws As Worksheet
'    On Error Resume Next
'    Set ws = ThisWorkbook.Sheets("Admin")
'
'    ' Check if the sheet exists and C4 has a value
'    If Not ws Is Nothing Then
'        GetAdminPassword = Trim(ws.Range("C4").value) ' Read password from C4
'    Else
'        GetAdminPassword = "" ' Return empty if sheet not found
'    End If
'
'    On Error GoTo 0
'End Function

Function IsAlpha(ByVal str As String) As Boolean
    ' This function checks if the input string is a single alphabetic character
    If Len(str) = 1 And (str Like "[A-Za-z]") Then
        IsAlpha = True
    Else
        IsAlpha = False
    End If
End Function

Public Function ExportNewLabelPrints()
    ' Standard Module - Appends new records from tblLabelPrints to Label_Prints.csv based on unique Print ID (column 19)

    ' --- Declaration of object variables ---
    Dim dictCSVIDs As Object
    Dim tbl As ListObject
    Dim ws As Worksheet

    ' --- Declaration of variant variables ---
    Dim arrLine As Variant
    Dim vData As Variant

    ' --- Declaration of string variables ---
    Dim strCSVPath As String
    Dim textLine As String
    Dim csvLine As String
    Dim cleanID As String
    Dim excelID As String

    ' --- Declaration of integer/long variables ---
    Dim ffIn As Integer
    Dim ffOut As Integer
    Dim i As Long
    Dim j As Long

    On Error GoTo Cleanup

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Path to the CSV file
    strCSVPath = GetWorkbookTableFilePath(FILE_LABEL_DETAILS)

    Set dictCSVIDs = CreateObject("Scripting.Dictionary")

    ' Load existing Print IDs from CSV
    ffIn = FreeFile
    Open strCSVPath For Input As #ffIn

    ' Skip header
    Line Input #ffIn, textLine

    Do While Not EOF(ffIn)
        Line Input #ffIn, textLine
        If Trim(textLine) <> "" Then
            arrLine = Split(textLine, ",")
            If UBound(arrLine) >= 18 Then
                cleanID = Trim(Replace(arrLine(18), """", "")) ' Clean CSV Print ID
                If Not dictCSVIDs.Exists(cleanID) Then
                    dictCSVIDs.Add cleanID, True
                End If
            End If
        End If
    Loop
    Close #ffIn

    ' Get source table
    Set ws = ThisWorkbook.Worksheets("Label Prints")
    Set tbl = ws.ListObjects("tblLabelPrints")
    vData = tbl.DataBodyRange.value

    ' Append new entries
    ffOut = FreeFile
    Open strCSVPath For Append As #ffOut

    For i = 1 To UBound(vData, 1)
        excelID = Trim(CStr(vData(i, 19))) ' Clean Excel Print ID
        If Not dictCSVIDs.Exists(excelID) Then
            csvLine = ""
            For j = 1 To UBound(vData, 2)
                If j = 1 Then
                    csvLine = """" & vData(i, j) & """"
                Else
                    csvLine = csvLine & ",""" & vData(i, j) & """"
                End If
            Next j
            Print #ffOut, csvLine
        End If
    Next i

    Close #ffOut

    Debug.Print "Checking PrintID: " & Trim(CStr(vData(i, 19)))
    Debug.Print "Existing CSV PrintID: " & cleanID

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Set dictCSVIDs = Nothing
    Set tbl = Nothing
    Set ws = Nothing
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

Public Sub QuickSort2D(ByRef arr As Variant, ByVal leftIndex As Long, ByVal rightIndex As Long, ByVal sortCol As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim temp As Variant
    Dim k As Long
    Dim colCount As Long

    If leftIndex >= rightIndex Then Exit Sub

    colCount = UBound(arr, 2)
    i = leftIndex
    j = rightIndex
    pivot = arr((leftIndex + rightIndex) \ 2, sortCol)

    Do While i <= j
        Do While arr(i, sortCol) < pivot
            i = i + 1
        Loop

        Do While arr(j, sortCol) > pivot
            j = j - 1
        Loop

        If i <= j Then
            For k = 1 To colCount
                temp = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = temp
            Next k
            i = i + 1
            j = j - 1
        End If
    Loop

    If leftIndex < j Then QuickSort2D arr, leftIndex, j, sortCol
    If i < rightIndex Then QuickSort2D arr, i, rightIndex, sortCol
End Sub

' Adds a timestamp to the ANSI formatted date
Public Function GetTransactionTimestamp() As String
    GetTransactionTimestamp = DateToANSI(Format(Now, "dd/mm/yyyy")) & " " & Format(Now, "hh:mm:ss")
End Function