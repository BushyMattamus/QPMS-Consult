Attribute VB_Name = "BusLogDataLoads"
Option Explicit

' Imports specified columns from QC_Number.csv into the "QC_Number" sheet
' Uses QueryTable (TEXT connection) to avoid opening the CSV as a workbook
' Reuses public helpers: FindHeaderCol, RM_Pad4, MakeTempSheetName, SheetExists

Public Sub LoadQCNumberData()
    Dim qcCsvPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long
    Dim r As Long
    Dim s As String

    On Error GoTo ErrorHandler

    ' Resolve source CSV path and destination sheet
    qcCsvPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    Set wsDest = ThisWorkbook.Worksheets("QC_Number")

    ' Clear existing contents and import whole CSV into the sheet
    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(qcCsvPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 105, , ERROR_CSV_IMPORT & qcCsvPath
    End If

     lastRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row
    
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
    
    ' --- Update named ranges ---
    If lastRow >= 2 Then
        ThisWorkbook.Names("NRqcNumStdName").RefersTo = wsDest.Range("C2:C" & lastRow)
    End If
    

    GoTo CleanExit

ErrorHandler:
    MsgBox "LoadQCNumberData error: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Set wsDest = Nothing
    Exit Sub
End Sub

Public Sub LoadRMMasterData()
    On Error GoTo ErrorHandler

    Dim rmCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long
    Dim sourceFolderPath As String
    Dim r As Long
    Dim s As String

    sourceFolderPath = MASTER_DATA_PATH & FOLDER_TABLES
    rmCSVPath = FILE_RM_MASTER
    Set wsDest = ThisWorkbook.Worksheets("RM_Master")

    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(rmCSVPath, wsDest, wsDest.Range("A1"), True, sourceFolderPath)
    If Not importSuccess Then
        Err.Raise vbObjectError + 105, , ERROR_CSV_IMPORT & rmCSVPath
    End If

    wsDest.Columns(1).NumberFormat = "@"
    lastRow = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row
    If lastRow >= 2 Then
        For r = 2 To lastRow
            s = CStr(wsDest.Cells(r, 1).value)
            s = Trim$(s)
            If Len(s) = 0 Then
                wsDest.Cells(r, 1).value = ""
            ElseIf Len(s) >= 4 Then
                wsDest.Cells(r, 1).value = s
            Else
                wsDest.Cells(r, 1).value = Right$("0000" & s, 4)
            End If
        Next r
    End If

    GoTo CleanExit

ErrorHandler:
    MsgBox "LoadRMMasterData error: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Set wsDest = Nothing
    Exit Sub
End Sub