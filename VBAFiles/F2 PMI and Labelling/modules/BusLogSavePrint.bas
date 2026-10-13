Attribute VB_Name = "BusLogSavePrint"
Sub SavePrintLabel(Optional ByVal isSampleView As Boolean = False)
    ' Batch Entry - Saves data, generates Word labels, prints, and saves PDF
    ' Uses PMI Number to locate label template and archives PDFs to ARCHIVE_PRODUCT_LABELS_PATH
    ' When isSampleView = True, displays 10 sample labels without printing or saving

    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim wsCurrent As Worksheet
    Dim wsLabelRegister As Worksheet
    Dim wsAveryLabels As Worksheet
    Dim wsFGMaster As Worksheet
    Dim tblLabelReg As ListObject
    Dim tblAvery As ListObject
    Dim labelRow As ListRow
    Dim averyRow As ListRow
    Dim fgRow As Range
    Dim batchNumber As String
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim expiryDate As String
    Dim qtyValue As String
    Dim mLotValue As String
    Dim mLot2Value As String
    Dim commentsValue As String
    Dim userName As String
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim archiveFolderPath As String
    Dim pdfFilename As String
    Dim pdfFullPath As String
    Dim averyLabelCode As String
    Dim productID As String
    Dim productDescription As String
    Dim qidCode As String
    Dim sql As String
    Dim tablesPath As String
    Dim docText As String
    Dim pmiPattern As String
    Dim fullPmiText As String
    Dim extractedVersion As String
    Dim ch As String
    Dim nextFileName As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim docContent As Object
    Dim searchRange As Object
    Dim foundRange As Object
    Dim conn As Object
    Dim rs As Object
    Dim lblTable As Object
    Dim origTable As Object
    Dim cleanTableRange As Object
    Dim cell1Copy As Object
    Dim cell2Copy As Object
    Dim rng As Object
    Dim rngSub As Object
    Dim printID As String
    Dim currentYear As String
    Dim maxSeq As Integer
    Dim entry As ListRow
    Dim newRow As ListRow
    Dim quantity As Double
    Dim labelCount As Integer
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim fileCount As Integer
    Dim labelsPerSheet As Double
    Dim r As Long
    Dim c As Long
    Dim tblIndex As Long
    Dim totalPages As Long
    Dim lastRow As Long
    Dim pmiCol As Long
    Dim quickCol As Long
    Dim statusCol As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim dotPos As Long
    Dim versionEndPos As Long
    Dim i As Long

    AppStateUtils.DisableAppUpdates

    ' --- Get worksheet references ---
    Set wsCurrent = ThisWorkbook.ActiveSheet
    Set wsLabelRegister = ThisWorkbook.Worksheets("Label Register")
    Set wsAveryLabels = ThisWorkbook.Worksheets("Avery Labels")
    Set wsFGMaster = ThisWorkbook.Worksheets("FG_Master")

    ' --- Retrieve field values using GetFieldValByLabelC and GetFieldValByLabelH ---
    batchNumber = GetFieldValByLabelH("Batch Number:")
    pmiNumber = GetFieldValByLabelH("PMI Number:")
    expiryDate = GetFieldValByLabelH("Expiry Date:")
    productDescription = GetFieldValByLabelH("Description:")

    If isSampleView Then
        qtyValue = "10"
    Else
        qtyValue = GetFieldValByLabelC("Quantity:")
    End If
    
    ' --- Convert expiry date to DD/MM/YY ---
    If IsDate(expiryDate) Then
        expiryDate = Format(CDate(expiryDate), "DD/MM/YY")
    End If
    
    If Not isSampleView Then
        If Trim(qtyValue) = "" Then
            MsgBox "Quantity must be a number greater than zero.", vbExclamation, "Validation Error"
            GoTo Cleanup
        End If
    End If
    
    quantity = CDbl(qtyValue)
    mLotValue = GetFieldValByLabelC("Mfg Lot No:")
    mLot2Value = GetFieldValByLabelC("Mfg Lot No 2:")
    commentsValue = GetFieldValByLabelC("Comments:")
    userName = Environ("USERNAME")
    
    ' --- Validate no commas ---
    If Not ValidateNoCommas(mLotValue, "Mfg Lot No") Then GoTo CleanExit
    If Not ValidateNoCommas(mLot2Value, "Mfg Lot No 2") Then GoTo CleanExit
    If Not ValidateNoCommas(commentsValue, "Comments") Then GoTo CleanExit

    ' --- Validate critical fields ---
    If Trim(batchNumber) = "" Then
        MsgBox "Batch Number cannot be empty.", vbExclamation, "Validation Error"
        GoTo Cleanup
    End If

    If Trim(pmiNumber) = "" Then
        MsgBox "PMI Number cannot be empty.", vbExclamation, "Validation Error"
        GoTo Cleanup
    End If

    If Trim(expiryDate) = "" Then
        MsgBox "Expiry Date cannot be empty.", vbExclamation, "Validation Error"
        GoTo Cleanup
    End If

    If Not isSampleView Then
        If Not IsNumeric(quantity) Or quantity <= 0 Then
            MsgBox "Quantity must be a number greater than zero.", vbExclamation, "Validation Error"
            GoTo Cleanup
        End If

        ' --- Check if batch has already been printed ---
        tablesPath = GetWorkbookTablesPath()
        Set conn = OpenADOConnection(tablesPath)
        
        If conn Is Nothing Then
            MsgBox "Unable to connect to database to verify batch print status.", vbExclamation, "Save Print Label"
            GoTo Cleanup
        End If
        
        sql = "SELECT COUNT(*) FROM [" & FILE_LABEL_DETAILS & "] " & _
              "WHERE [" & LABEL_HEADER_BATCH_NUMBER & "] = '" & Replace(batchNumber, "'", "''") & "'"
        
        Set rs = conn.Execute(sql)
        
        If Not rs.EOF Then
            If rs.Fields(0).value > 0 Then
                MsgBox "This batch has already been printed. Please use the reprint facility.", vbExclamation, "Save Print Label"
                rs.Close
                conn.Close
                Set rs = Nothing
                Set conn = Nothing
                GoTo Cleanup
            End If
        End If
        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
    End If

    ' --- Lookup Avery Label Code in Label Register using PMI Number ---
    Set tblLabelReg = wsLabelRegister.ListObjects("tblLabelReg")
    averyLabelCode = ""
    productID = ""

    For Each labelRow In tblLabelReg.ListRows
        If labelRow.Range(1).value = pmiNumber Then
            averyLabelCode = labelRow.Range(2).Text
            Exit For
        End If
    Next labelRow

    If averyLabelCode = "" Then
        MsgBox "Missing Avery Label Code for PMI Number " & pmiNumber & " in Label Register.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' --- Lookup Quick Code from FG_Master using PMI Number ---
    qidCode = ""

    lastRow = wsFGMaster.Cells(wsFGMaster.Rows.Count, 1).End(xlUp).row
    pmiCol = 0
    quickCol = 0
    statusCol = 0

    For i = 1 To wsFGMaster.Cells(1, wsFGMaster.Columns.Count).End(xlToLeft).Column
        Select Case wsFGMaster.Cells(1, i).value
            Case "PMI_Code"
                pmiCol = i
            Case "Quick_Code"
                quickCol = i
            Case "Product_Status"
                statusCol = i
        End Select
        If pmiCol > 0 And quickCol > 0 And statusCol > 0 Then Exit For
    Next i

    If pmiCol > 0 And quickCol > 0 And statusCol > 0 Then
        For i = 2 To lastRow
            If wsFGMaster.Cells(i, pmiCol).value = pmiNumber And _
               wsFGMaster.Cells(i, statusCol).value = "Current" Then
                qidCode = wsFGMaster.Cells(i, quickCol).value
                Exit For
            End If
        Next i
    End If

    If qidCode = "" Then
        MsgBox "Quick Code not found for PMI Number " & pmiNumber & " with Product_Status = 'Current' in FG_Master.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' --- Lookup Labels Per Sheet in Avery Labels ---
    Set tblAvery = wsAveryLabels.ListObjects("tblAvery")
    labelsPerSheet = 0

    For Each averyRow In tblAvery.ListRows
        If averyRow.Range(1).value = averyLabelCode Then
            labelsPerSheet = averyRow.Range(4).value
            Exit For
        End If
    Next averyRow

    If labelsPerSheet = 0 Then
        MsgBox "Labels per Sheet for Avery Label Code " & averyLabelCode & " not found.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' --- Construct path to approved labels folder ---
    folderPath = GetFolderDrivePath(PRODUCT_APPROVED_LABELS_PATH)

    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Product Approved Labels folder not found: " & folderPath, vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' --- Validate only one label file exists for this PMI Number ---
    If Not ValidateUniqueTemplateFile(folderPath, pmiNumber & "-", pmiNumber & "-*.docx", "Label", fileName) Then
        GoTo Cleanup
    End If

    fullPath = folderPath & fileName

    If Not isSampleView Then
        ' --- Construct archive folder path ---
        archiveFolderPath = GetFolderDrivePath(ARCHIVE_PRODUCT_LABELS_PATH)

        If Not FolderExists(archiveFolderPath) Then
            MsgBox "Archive Labels folder not found: " & archiveFolderPath, vbExclamation, "Error"
            GoTo Cleanup
        End If

        ' --- Generate next Print ID ---
        printID = GetNextLabelPrintID()
    End If

    ' --- Open or get Word application ---
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo ErrHandler

    wdApp.Visible = True

    ' --- Bring Word window to front ---
    On Error Resume Next
    wdApp.Activate
    DoEvents
    On Error GoTo ErrHandler
    
    ' --- Open document (read-only for sample view) ---
    If isSampleView Then
        Set wdDoc = wdApp.Documents.Open(fullPath, ReadOnly:=True)
    Else
        Set wdDoc = wdApp.Documents.Open(fullPath)
    End If

    ' --- Retrieve PMI Version from form ---
    pmiVersion = GetFieldValByLabelH("PMI Version:")

    ' --- Extract and validate version from document ---
    If Not ValidateLabelPMIVersion(wdDoc, wdApp, pmiNumber, pmiVersion) Then
        GoTo Cleanup
    End If

    ' --- Validate manufacturing lot numbers if required by label template ---
    If Not ValidateLabelMfgLotNumbers(wdDoc, wdApp, mLotValue, mLot2Value) Then
        GoTo Cleanup
    End If

    ' --- Replace placeholders in document ---
    With wdDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = 0
        .MatchCase = False

        .Text = "?BATCH"
        .Replacement.Text = batchNumber
        .Execute Replace:=2

        .Text = "?EXP"
        .Replacement.Text = "Exp: " & expiryDate
        .Execute Replace:=2

        .Text = "?MLOT2"
        .Replacement.Text = mLot2Value
        .Execute Replace:=2

        .Text = "?MLOT"
        .Replacement.Text = mLotValue
        .Execute Replace:=2
    End With

    ' --- Bold "EXP: " text in document ---
    Set rng = wdDoc.Content

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "Exp: "
        .Forward = True
        .Wrap = 0
        .MatchCase = False
    End With

    Do While rng.Find.Execute
        rng.Font.Bold = True
        rng.Collapse Direction:=0
    Loop

    ' --- Adjust margins ---
    wdDoc.PageSetup.BottomMargin = wdApp.CentimetersToPoints(0.44)

    ' --- Setup label table replication ---
    Set cleanTableRange = wdDoc.Tables(1).Range.Duplicate
    Set lblTable = wdDoc.Tables(1)
    Set origTable = wdDoc.Tables(1)

    Set cell1Copy = origTable.Cell(1, 1).Range.Duplicate
    Set cell2Copy = origTable.Cell(1, 2).Range.Duplicate

    ' --- Replicate labels based on Avery Label Code ---
    If averyLabelCode = "L7158" Then

        ' --- First pass: Populate columns 1, 4, 7... ---
        cell1Copy.Copy
        labelCount = 1
        Do While labelCount < quantity
            For rowIndex = 1 To lblTable.Rows.Count
                For colIndex = 1 To lblTable.Columns.Count Step 3
                    If Not (rowIndex = 1 And colIndex = 1 And labelCount = 1) Then
                        If labelCount >= quantity Then Exit For
                        lblTable.Cell(rowIndex, colIndex).Range.Paste
                        labelCount = labelCount + 1
                    End If
                Next colIndex
                If labelCount >= quantity Then Exit For
            Next rowIndex

            If labelCount < quantity Then
                cleanTableRange.Copy
                With wdDoc.Range
                    .Collapse Direction:=0
                    .InsertBreak Type:=7
                    .Paste
                End With
                Set lblTable = wdDoc.Tables(wdDoc.Tables.Count)
                For r = 1 To lblTable.Rows.Count
                    For c = 1 To lblTable.Columns.Count
                        lblTable.Cell(r, c).Range.Text = ""
                    Next c
                Next r
            End If
        Loop

        ' --- Second pass: Populate columns 2, 5, 8... across all tables ---
        Set cell2Copy = wdDoc.Tables(1).Cell(1, 2).Range.Duplicate
        cell2Copy.Copy

        labelCount = 0
        totalPages = wdDoc.Tables.Count

        For tblIndex = 1 To totalPages
            Set lblTable = wdDoc.Tables(tblIndex)

            For rowIndex = 1 To lblTable.Rows.Count
                For colIndex = 2 To lblTable.Columns.Count Step 3
                    If labelCount >= quantity Then Exit For
                    lblTable.Cell(rowIndex, colIndex).Range.Paste
                    labelCount = labelCount + 1
                Next colIndex
                If labelCount >= quantity Then Exit For
            Next rowIndex

            If labelCount >= quantity Then Exit For
        Next tblIndex

    Else

        ' --- Default: Populate columns 1, 3, 5... ---
        cell1Copy.Copy
        labelCount = 1
        Do While labelCount < quantity
            For rowIndex = 1 To lblTable.Rows.Count
                For colIndex = 1 To lblTable.Columns.Count Step 2
                    If Not (rowIndex = 1 And colIndex = 1 And labelCount = 1) Then
                        If labelCount >= quantity Then Exit For
                        On Error Resume Next
                        lblTable.Cell(rowIndex, colIndex).Range.Paste
                        On Error GoTo ErrHandler
                        labelCount = labelCount + 1
                    End If
                Next colIndex
                If labelCount >= quantity Then Exit For
            Next rowIndex

            If labelCount < quantity Then
                cleanTableRange.Copy
                With wdDoc.Range
                    .Collapse Direction:=0
                    .InsertBreak Type:=7
                    .Paste
                End With
                Set lblTable = wdDoc.Tables(wdDoc.Tables.Count)
                For r = 1 To lblTable.Rows.Count
                    For c = 1 To lblTable.Columns.Count
                        lblTable.Cell(r, c).Range.Text = ""
                    Next c
                Next r
            End If
        Loop

    End If

    ' --- Update fields ---
    wdDoc.Fields.Update

    ' --- For sample view, exit here leaving document open for review ---
    If isSampleView Then
        GoTo CleanExit
    End If

    ' --- Print document ---
    wdDoc.PrintOut

    ' --- Build PDF filename and path ---
    pdfFilename = pmiNumber & "_" & batchNumber & "_" & printID & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
    pdfFullPath = archiveFolderPath & pdfFilename

    ' --- Export as PDF ---
    wdDoc.ExportAsFixedFormat OutputFileName:=pdfFullPath, ExportFormat:=17

    ' --- Activate Word and close ---
    wdApp.Activate
    wdDoc.Close SaveChanges:=False
    ' If wdApp.Documents.Count = 0 Then wdApp.Quit
    wdApp.Quit

    expiryDate = Format(expiryDate, "yyyy-mm-dd")

    ' --- Write print record to Label_Prints.csv ---
    If Not AppendLabelPrintRecord( _
        batchNumber:=batchNumber, _
        quantity:=CStr(quantity), _
        pmiNumber:=pmiNumber, _
        expiryDate:=expiryDate, _
        qidCode:=qidCode, _
        mfgLotNo:=mLotValue, _
        mfgLotNo2:=mLot2Value, _
        averyLabelCode:=averyLabelCode, _
        labelsPerSheet:=CStr(labelsPerSheet), _
        productID:="productID", _
        cpmuRef:=pmiNumber, _
        productDescription:=productDescription, _
        printType:="Original", _
        comments:=commentsValue) Then
        
        MsgBox "Warning: Failed to write print record to Label_Prints.csv", vbExclamation, "SavePrintLabel"
    End If
    
    BEClrNew

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If

CleanExit:
    On Error Resume Next
    Set wsFGMaster = Nothing
    Set wsAveryLabels = Nothing
    Set wsLabelRegister = Nothing
    Set wsCurrent = Nothing
    Set rng = Nothing
    If Not isSampleView Then
        Set wdDoc = Nothing
        Set wdApp = Nothing
    End If
    AppStateUtils.RestoreAppState
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "SavePrintLabel"
    Resume Cleanup

End Sub
