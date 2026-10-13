Attribute VB_Name = "BusLogRecordLoader"
Public Sub ResolveRMCodeFromDescription(sourceCell As Range, outputCell As Range)
    Dim desc As String
    desc = Trim$(CStr(sourceCell.value))
    If desc = "" Then Exit Sub

    Dim wsRM As Worksheet
    Set wsRM = ThisWorkbook.Worksheets("RM_Master")
    
    Dim colStd As Long, colRM As Long
    colStd = RM_COL_STDNAME
    colRM = RM_COL_RMCODE
    
    If colStd = 0 Or colRM = 0 Then
        MsgBox "RM Master columns not initialized. Please restart the workbook.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long, r As Long
    lastRow = wsRM.Cells(wsRM.Rows.Count, colStd).End(xlUp).row
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim valStd As String, valRM As String
    For r = 2 To lastRow
        valStd = Trim$(CStr(wsRM.Cells(r, colStd).value))
        If StrComp(valStd, desc, vbTextCompare) = 0 Then
            valRM = Trim$(CStr(wsRM.Cells(r, colRM).value))
            If valRM <> "" Then
                If Not dict.Exists(valRM) Then dict.Add valRM, 1
            End If
        End If
    Next r
    
    If dict.Count = 1 Then
        outputCell.NumberFormat = "@"
        outputCell.value = RM_Pad4(dict.Keys()(0))
    Else
        outputCell.value = "Select from list"
    End If

    GoTo CleanExit

ErrorHandler:
    MsgBox "ResolveRMCodeFromDescription error: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not dict Is Nothing Then Set dict = Nothing
    Set wsRM = Nothing
    Exit Sub
    
End Sub