Attribute VB_Name = "BusLogDataLoads"
Option Explicit

' Imports specified columns from QC_Number.csv into the "QC_Number" sheet
' Uses QueryTable (TEXT connection) to avoid opening the CSV as a workbook
' Reuses public helpers: FindHeaderCol, RM_Pad4, MakeTempSheetName, SheetExists

Public Sub LoadQCNumberData()
    Dim qcCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long
    Dim r As Long
    Dim s As String

    On Error GoTo ErrorHandler
    
    ' Resolve source CSV path and destination sheet
    qcCSVPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    Set wsDest = ThisWorkbook.Worksheets("QC_Number")

    ' Clear existing contents and import whole CSV into the sheet
    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(qcCSVPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 105, , ERROR_CSV_IMPORT & qcCSVPath
    End If

    ' Ensure column B is stored as text and formatted/padded to 4 characters ("0000")
    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    
    If lastRow >= 2 Then
        wsDest.Columns(1).NumberFormat = "@"
        For r = 2 To lastRow
            s = CStr(wsDest.Cells(r, 1).value)
            s = Trim$(s)
            If Len(s) = 0 Then
                wsDest.Cells(r, 1).value = ""
            ElseIf Len(s) >= 6 Then
                wsDest.Cells(r, 1).value = s
            Else
                wsDest.Cells(r, 1).value = Right$("000000" & s, 6)
            End If
        Next r
    End If

    wsDest.Columns(2).NumberFormat = "@"
    If lastRow >= 2 Then
        For r = 2 To lastRow
            s = CStr(wsDest.Cells(r, 2).value)
            s = Trim$(s)
            If Len(s) = 0 Then
                wsDest.Cells(r, 2).value = ""
            ElseIf Len(s) >= 4 Then
                wsDest.Cells(r, 2).value = s
            Else
                wsDest.Cells(r, 2).value = Right$("0000" & s, 4)
            End If
        Next r
    End If

    ' Keep sheet hidden as before
    ' wsDest.Visible = xlSheetHidden
    
    Exit Sub
    
ErrorHandler:
    MsgBox "LoadQCNumberData error: " & Err.Description, vbCritical
    Exit Sub

End Sub