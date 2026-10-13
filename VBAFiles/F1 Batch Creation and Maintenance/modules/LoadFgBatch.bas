Attribute VB_Name = "LoadFgBatch"
' filepath: F1 Finished Goods Batch\LoadFgBatch.bas

Sub ImportFgBatchSelectedCSVColumns()

    Dim csvPath As String
    Dim tempCsvPath As String
    Dim destSheetName As String: destSheetName = "BatchRef"
    Dim wsDest As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim lines As Variant
    Dim line As String
    Dim i As Long
    Dim j As Long
    Dim fields As Variant
    Dim colsToExtract As Variant
    Dim dict As Object
    Dim uniqueKey As String
    Dim outputRow As Long
    Dim fileContent As String

    On Error GoTo ErrHandler

    ' --- Create temporary copy of CSV ---
    tempCsvPath = CreateTemporaryFileCopy(FILE_FINISHED_GOODS_BATCH_NUMBER_MASTER, PATH_MATERIAL_MANAGEMENT_TABLES)
    If tempCsvPath = "" Then
        MsgBox "Unable to create temporary copy of source file.", vbExclamation, "Import Error"
        GoTo CleanExit
    End If

    ' Columns A–K (0–10)
    colsToExtract = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    ' --- Read from temporary CSV ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(tempCsvPath, 1, False)

    fileContent = ts.ReadAll
    ts.Close

    ' --- Remove UTF-8 BOM if present (TEMPORARY - should not occur with ANSI files) ---
    If Left(fileContent, 3) = Chr(239) & Chr(187) & Chr(191) Then
        Debug.Print "WARNING: UTF-8 BOM found in " & tempCsvPath & " - file should be ANSI"
        fileContent = Mid(fileContent, 4)
    End If

    lines = Split(fileContent, vbCrLf)

    ' Remove trailing blanks
    Do While UBound(lines) >= 0 And Trim(lines(UBound(lines))) = ""
        ReDim Preserve lines(0 To UBound(lines) - 1)
    Loop

    If UBound(lines) < 0 Then GoTo CleanExit

    ' -- Dedup dictionary --
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Dim result()
    ReDim result(0 To UBound(lines), 0 To UBound(colsToExtract))
    outputRow = 0

    ' --- Parse CSV ---
    For i = 0 To UBound(lines)

        fields = ParseCSVLine(CStr(lines(i)))

        ' Build dedupe key
        uniqueKey = ""
        For j = 0 To UBound(colsToExtract)
            uniqueKey = uniqueKey & "|" & CStr(fields(colsToExtract(j)))
        Next j

        ' Add unique rows only
        If i = 0 Or Not dict.Exists(uniqueKey) Then
            dict(uniqueKey) = True

            For j = 0 To UBound(colsToExtract)
                Dim val As Variant
                val = fields(colsToExtract(j))

                ' Column 8 (index 7) MUST stay as TEXT (DD/MM/YYYY)
                If colsToExtract(j) = 7 Then
                    val = CStr(val)                ' FORCE STRING
                End If

                result(outputRow, j) = val
            Next j

            outputRow = outputRow + 1
        End If
    Next i

    ' --- Destination Sheet ---
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    On Error GoTo ErrHandler

    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add
        wsDest.Name = destSheetName
    Else
        wsDest.Range("A:M").ClearContents
    End If

    ' Pre-format columns as TEXT BEFORE writing the array.
    wsDest.Range("A:M").NumberFormat = "@"

    ' --- Write data to sheet ---
    If outputRow > 0 Then
        wsDest.Range(wsDest.Cells(1, 1), _
                     wsDest.Cells(outputRow, UBound(colsToExtract) + 1)).value = result
    End If
    
    ' --- Apply Excel SORT formulas in columns N and O ---
     If outputRow > 1 Then
        Dim dataEndRow As Long
        dataEndRow = outputRow

        ' Headers
        wsDest.Range("N1").value = "Sorted A"
        
        ' Unique + Sorted column A (col 1)
        wsDest.Range("N2").Formula2 = _
            "=SORT(UNIQUE(FILTER(A2:A" & dataEndRow & ",A2:A" & dataEndRow & "<>"""")))"
        
     End If

    ' --- Set sheet visibility based on environment ---
    If PROD_DEV <> "DEV" Then
        wsDest.Visible = xlSheetHidden
    End If

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
    Exit Sub

ErrHandler:
    MsgBox "Error importing CSV: " & Err.Description, vbCritical
    Resume CleanExit

End Sub
