Attribute VB_Name = "LoadFgMaster"
' ========================================================================
' FINISHED GOODS MASTER DATA IMPORT
' ========================================================================
' Imports selected columns from the Finished Goods Master CSV file,
' removes duplicates, filters archived/reserved records, and populates
' the reference sheet with unique sorted data.
' ========================================================================

Sub ImportSelectedCSVColumns()
    On Error GoTo ErrHandler

    Dim tempCsvPath As String
    Dim csvPath As String
    Dim destSheetName As String
    Dim wsDest As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim lines As Variant
    Dim line As String
    Dim i As Long
    Dim j As Long
    Dim sourceData() As Variant
    Dim resultData() As Variant
    Dim colsToExtract As Variant
    Dim rowCount As Long
    Dim dict As Object
    Dim uniqueKey As String
    Dim outputRow As Long
    Dim fields As Variant
    Dim statusVal As String
    Dim fileContent As String
    Dim finalData() As Variant
    Dim dataEndRow As Long

    ' --- Create temporary copy of CSV from Master Data location ---
    tempCsvPath = CreateTemporaryFileCopy(FILE_MASTER_DATA_FINISHED_GOODS, GetMasterDataTablesPath())
    If tempCsvPath = "" Then
        MsgBox "Unable to create temporary copy of source file.", vbExclamation, "Import Error"
        GoTo CleanExit
    End If

    ' --- Columns to extract: A, B, F, G, H, I (0-based: 0, 1, 5, 6, 7, 8) ---
    colsToExtract = Array(0, 1, 5, 6, 7, 8)
    destSheetName = "MasterRef"

    ' --- Read CSV into array using FileSystemObject ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(tempCsvPath, 1, False)

    ' --- Read all lines ---
    fileContent = ts.ReadAll
    ts.Close

    ' --- Remove UTF-8 BOM if present (TEMPORARY - should not occur with ANSI files) ---
    If Left(fileContent, 3) = Chr(239) & Chr(187) & Chr(191) Then
        Debug.Print "WARNING: UTF-8 BOM found in " & tempCsvPath & " - file should be ANSI"
        fileContent = Mid(fileContent, 4)
    End If

    lines = Split(fileContent, vbCrLf)

    ' --- Remove trailing empty lines ---
    Do While UBound(lines) >= 0 And Trim(lines(UBound(lines))) = ""
        ReDim Preserve lines(0 To UBound(lines) - 1)
    Loop
    If UBound(lines) < 0 Then
        GoTo CleanExit
    End If

    rowCount = UBound(lines) + 1

    ' --- Parse CSV lines into 2D array ---
    ReDim sourceData(0 To rowCount - 1, 0 To 19)
    For i = 0 To UBound(lines)
        fields = ParseCSVLine(CStr(lines(i)))
        For j = 0 To UBound(fields)
            If j <= 19 Then
                sourceData(i, j) = fields(j)
            End If
        Next j
    Next i

    ' --- Extract desired columns and remove duplicates ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    ReDim resultData(0 To rowCount - 1, 0 To UBound(colsToExtract))
    outputRow = 0

    For i = 0 To UBound(sourceData, 1)
        ' --- Skip Archived or Reserved rows ---
        statusVal = LCase(Trim(CStr(sourceData(i, 6))))

        If statusVal = "archived" Or statusVal = "reserved" Then
            GoTo SkipRow
        End If

        uniqueKey = ""
        For j = 0 To UBound(colsToExtract)
            uniqueKey = uniqueKey & "|" & CStr(sourceData(i, colsToExtract(j)))
        Next j

        If i = 0 Or Not dict.Exists(uniqueKey) Then
            For j = 0 To UBound(colsToExtract)
                resultData(outputRow, j) = sourceData(i, colsToExtract(j))
            Next j
            If i > 0 Then
                dict.Add uniqueKey, True
            End If
            outputRow = outputRow + 1
        End If

SkipRow:
    Next i

    ' --- Reference sheet handling ---
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    On Error GoTo ErrHandler

    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add
        wsDest.Name = destSheetName
    Else
        wsDest.Range("A:H").Clear
    End If

    ' --- Write array to sheet in one operation ---
    If outputRow > 0 Then
        ReDim finalData(0 To outputRow - 1, 0 To UBound(colsToExtract))

        For i = 0 To outputRow - 1
            For j = 0 To UBound(colsToExtract)
                finalData(i, j) = resultData(i, j)
            Next j
        Next i

        wsDest.Range(wsDest.Cells(1, 1), _
                     wsDest.Cells(outputRow, UBound(colsToExtract) + 1)).value = finalData
    End If

    ' --- Apply Excel SORT formulas in columns N and O ---
    If outputRow > 1 Then
        dataEndRow = outputRow

        wsDest.Range("N1").value = "Sorted A"
        wsDest.Range("O1").value = "Sorted E"

        wsDest.Range("N2").Formula2 = _
            "=SORT(UNIQUE(FILTER(A2:A" & dataEndRow & ",A2:A" & dataEndRow & "<>"""")))"

        wsDest.Range("O2").Formula2 = _
            "=SORT(UNIQUE(FILTER(C2:C" & dataEndRow & ",C2:C" & dataEndRow & "<>"""")))"
    End If

    ' --- Set sheet visibility based on environment ---
    If PROD_DEV <> "DEV" Then
        wsDest.Visible = xlSheetHidden
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error in ImportSelectedCSVColumns: " & Err.Description, vbCritical
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then
        ts.Close
    End If
    If Not fso Is Nothing And tempCsvPath <> "" Then
        If fso.FileExists(tempCsvPath) Then
            fso.DeleteFile tempCsvPath
        End If
    End If
    On Error GoTo 0
    Set fso = Nothing
    Set ts = Nothing
    Set dict = Nothing
    Set wsDest = Nothing
    Exit Sub
End Sub
