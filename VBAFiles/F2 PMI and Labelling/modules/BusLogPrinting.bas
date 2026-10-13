Attribute VB_Name = "BusLogPrinting"
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal uFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal uFlags As Long) As Long
#End If

Const HWND_TOPMOST As Long = -1
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Long = &H1
Const SWP_SHOWWINDOW As Long = &H40

Sub PRPrint()
    On Error GoTo ErrHandler

    Dim wsPrints As Worksheet
    Dim wsLabelPrints As Worksheet
    Dim wsReprintReasons As Worksheet
    Dim tblPrints As ListObject
    Dim tblLabelPrintsImp As ListObject
    Dim tblReprintReasons As ListObject
    Dim printRow As ListRow
    Dim labelPrintRow As ListRow
    Dim reprintReason As ListRow
    Dim labelsToPrint As Variant
    Dim printID As String
    Dim qidCode As String
    Dim batchNumber As String
    Dim printStatus As String
    Dim averyLabelCode As String
    Dim printDate As String
    Dim printTime As String
    Dim pdfFilename As String
    Dim pdfFullPath As String
    Dim pdfFolderPath As String
    Dim fileName As String
    Dim folderPath As String
    Dim productDescription As String
    Dim createdBy As String
    Dim productID As String
    Dim reprintID As String
    Dim expiryDate As String
    Dim mLot As String
    Dim mLot2 As String
    Dim reason As String
    Dim userName As String
    Dim pmiNumber As String
    Dim quantity As Integer
    Dim labelsPerSheet As Integer
    Dim labelsRequested As Integer
    Dim labelCount As Integer
    Dim createdDate As Date
    Dim reasonMatch As Boolean
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim hWordWnd As LongPtr
    Dim lblTable As Object
    Dim origTable As Object
    Dim cleanTableRange As Object
    Dim cell1Copy As Object
    Dim cell2Copy As Object
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim r As Long
    Dim c As Long
    Dim tblIndex As Long
    Dim totalPages As Long
    Dim rng As Object
    Dim rngSub As Object
    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim tempFolderPath As String
    Dim tempFileName As String
    Dim sql As String
    Dim fso As Object

    Set wsPrints = ThisWorkbook.Worksheets("Reprint Labels")
    Set wsLabelPrints = ThisWorkbook.Worksheets("Label_Prints")
    Set wsReprintReasons = ThisWorkbook.Worksheets("Reprint Reasons")
    Set tblPrints = wsPrints.ListObjects("tblPrints")
    Set tblLabelPrintsImp = wsLabelPrints.ListObjects("tblLabelPrintsImp")
    Set tblReprintReasons = wsReprintReasons.ListObjects(1)
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    AppStateUtils.DisableAppUpdates

    For Each printRow In tblPrints.ListRows

        If UCase(printRow.Range.Cells(1, 11).value) = "X" Then

            printID = printRow.Range.Cells(1, 1).Text
            qidCode = printRow.Range.Cells(1, 2).Text
            batchNumber = printRow.Range.Cells(1, 5).Text
            mLot = printRow.Range.Cells(1, 6).Text
            mLot2 = printRow.Range.Cells(1, 7).Text
            labelsPerSheet = printRow.Range.Cells(1, 9).value
            labelsRequested = printRow.Range.Cells(1, 10).value
            labelsToPrint = printRow.Range.Cells(1, 12).value
            reason = Trim(printRow.Range.Cells(1, 13).Text)
            userName = Environ("UserName")
            
            ' --- Retrieve batch data from FinishedGoodsBatchNumberMasterFile ---
            tempPath = CreateTemporaryFileCopy(FILE_FINISHED_GOODS_BATCH_MASTER)
            
            If tempPath = "" Then
                MsgBox "Failed to create temporary copy of batch master file.", vbCritical, "PRPrint"
                GoTo Cleanup
            End If
            
            tempFolderPath = fso.GetParentFolderName(tempPath)
            tempFileName = fso.GetFileName(tempPath)
            
            Set conn = OpenADOConnection(tempFolderPath)
            
            If conn Is Nothing Then
                MsgBox "Failed to open database connection.", vbCritical, "PRPrint"
                GoTo Cleanup
            End If
            
            sql = "SELECT [" & FGBNM_HEADER_PMI_NUMBER & "], " & _
                  "[" & FGBNM_HEADER_PRODUCT_DESCRIPTION & "], " & _
                  "[" & FGBNM_HEADER_EXPIRY_DATE & "] " & _
                  "FROM [" & tempFileName & "] " & _
                  "WHERE [" & FGBNM_HEADER_BATCH_NUMBER & "] = '" & Replace(batchNumber, "'", "''") & "'"
            
            Set rs = conn.Execute(sql)
            
            If rs.EOF Then
                MsgBox "Batch number " & batchNumber & " not found in batch master file.", vbExclamation, "PRPrint"
                GoTo CleanupConnection
            End If
            
            pmiNumber = rs.Fields(FGBNM_HEADER_PMI_NUMBER).value & ""
            productDescription = rs.Fields(FGBNM_HEADER_PRODUCT_DESCRIPTION).value & ""
            expiryDate = rs.Fields(FGBNM_HEADER_EXPIRY_DATE).value & ""
            
            rs.Close
            Set rs = Nothing
            conn.Close
            Set conn = Nothing
            
            If fso.FileExists(tempPath) Then
                fso.DeleteFile tempPath, True
            End If
            
            ' --- Convert expiry date from source format to DD/MM/YY ---
            If IsDate(expiryDate) Then
                expiryDate = Format(CDate(expiryDate), "DD/MM/YY")
            End If
            
            For Each labelPrintRow In tblLabelPrintsImp.ListRows
                If labelPrintRow.Range(19) = printID Then
                    createdBy = labelPrintRow.Range(7).Text
                    createdDate = labelPrintRow.Range(8).Text
                    averyLabelCode = labelPrintRow.Range(9).Text
                    productID = labelPrintRow.Range(12).Text
                    printStatus = labelPrintRow.Range(15).Text
                    Exit For
                End If
            Next labelPrintRow

            If IsEmpty(labelsToPrint) Then
                quantity = labelsRequested
            Else
                quantity = labelsToPrint
            End If

            If reason = "" Or reason = vbNullString Then
                MsgBox "Please enter a Reprint Reason before proceeding with the print job.", vbExclamation, "Missing Reason"
                GoTo Cleanup
            End If

            printDate = Format(Now, "YYYYMMDD")
            printTime = Format(Now, "HHMMSS")

            If IsEmpty(labelsToPrint) And IsEmpty(labelsRequested) Then
                MsgBox "Error: Please enter the number of labels to print in the 'Number of Prints' column."
                GoTo Cleanup
            End If

            folderPath = GetFolderDrivePath(PRODUCT_APPROVED_LABELS_PATH)

            pdfFolderPath = GetFolderDrivePath(ARCHIVE_PRODUCT_LABELS_PATH)

            fileName = Dir(folderPath & pmiNumber & "*.docx")
            If fileName = "" Then
                MsgBox "Template file for PMI Number " & pmiNumber & " not found.", vbExclamation, "Error"
                GoTo Cleanup
            End If

            Dim fileCount As Integer
            Dim nextFileName As String
            fileCount = 1
            nextFileName = Dir()

            Do While nextFileName <> ""
                If nextFileName Like pmiNumber & "*.docx" Then
                    fileCount = fileCount + 1
                End If
                nextFileName = Dir()
            Loop

            If fileCount > 1 Then
                MsgBox "Multiple label template files found for PMI Number: " & pmiNumber & vbCrLf & _
                       "Expected exactly one file starting with '" & pmiNumber & "*.docx'" & vbCrLf & _
                       "Please remove duplicate label templates from the Product Approved Labels folder.", vbExclamation, "Validation Error"
                GoTo Cleanup
            End If

            Set wdApp = CreateObject("Word.Application")
            wdApp.Visible = True

            hWordWnd = FindWindow("OpusApp", wdApp.Caption)
            If hWordWnd <> 0 Then
                SetWindowPos hWordWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
            End If

            Set wdDoc = wdApp.Documents.Open(folderPath & fileName)

            With wdDoc
                .Fields.Update
                .Content.Find.Execute FindText:="?BATCH", ReplaceWith:=batchNumber, Replace:=2
                .Content.Find.Execute FindText:="?EXP", ReplaceWith:="Exp: " & expiryDate, Replace:=2
                .Content.Find.Execute FindText:="?MLOT2", ReplaceWith:=mLot2, Replace:=2
                .Content.Find.Execute FindText:="?MLOT", ReplaceWith:=mLot, Replace:=2

                Set rng = wdDoc.Content

                With rng.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = "Exp:"
                    .Forward = True
                    .Wrap = wdFindStop
                End With

                Do While rng.Find.Execute
                    If (rng.End - rng.Start) >= 3 Then
                        Set rngSub = rng.Duplicate
                        rngSub.End = rngSub.Start + 3
                        rngSub.Font.Bold = True
                    End If
    
                    ' Collapse the range to continue searching beyond this instance
                    rng.Collapse Direction:=wdCollapseEnd
                Loop

                .PageSetup.BottomMargin = wdApp.CentimetersToPoints(0.44)

                Set cleanTableRange = wdDoc.Tables(1).Range.Duplicate
                Set lblTable = wdDoc.Tables(1)
                Set origTable = wdDoc.Tables(1)

                Set cell1Copy = origTable.Cell(1, 1).Range.Duplicate
                Set cell2Copy = origTable.Cell(1, 2).Range.Duplicate

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

                    Set cell2Copy = wdDoc.Tables(1).Cell(1, 2).Range.Duplicate
                    cell2Copy.Copy

                    ' start counter at zero so the very first paste goes to (1,2)
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
                    ' Loop through rows and only columns 1, 3, 5, creating multiple tables if necessary
                    Set cell1Copy = lblTable.Cell(1, 1).Range.Duplicate
                    cell1Copy.Copy
                    labelCount = 1 ' we already populated the first label manually
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

                        ' If more labels needed, duplicate the table and clear its contents
                        If labelCount < quantity Then
                            cleanTableRange.Copy
                            With wdDoc.Range
                                .Collapse Direction:=0
                                .InsertBreak Type:=7
                                .Paste
                            End With
                            Set lblTable = wdDoc.Tables(wdDoc.Tables.Count)

                            ' Clear all cell contents in the pasted table
                            For r = 1 To lblTable.Rows.Count
                                For c = 1 To lblTable.Columns.Count
                                    lblTable.Cell(r, c).Range.Text = ""
                                Next c
                            Next r
                        End If
                    Loop
                End If

                With wdDoc.PageSetup
                    .FirstPageTray = wdPrinterDefaultBin
                    .OtherPagesTray = wdPrinterDefaultBin
                End With

                .PrintOut

                reprintID = GetNextLabelPrintID()

                pdfFilename = pmiNumber & "_" & batchNumber & "_" & reprintID & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
                pdfFullPath = pdfFolderPath & pdfFilename

                If Dir(pdfFolderPath, vbDirectory) = "" Then
                    MsgBox "PDF folder path does not exist: " & pdfFolderPath, vbExclamation, "Error"
                    GoTo Cleanup
                End If

                .ExportAsFixedFormat OutputFileName:=pdfFullPath, ExportFormat:=17

            End With

            wdDoc.Close False
            wdApp.Quit

            reasonMatch = False

            For Each reprintReason In tblReprintReasons.ListRows
                If UCase(reason) = UCase(reprintReason.Range.Cells(1, 1).value) Then
                    reasonMatch = True
                    Exit For
                End If
            Next reprintReason
            
            expiryDate = Format(expiryDate, "yyyy-mm-dd")

            If Not AppendLabelPrintRecord( _
                batchNumber:=batchNumber, _
                quantity:=CStr(quantity), _
                pmiNumber:=pmiNumber, _
                expiryDate:=expiryDate, _
                qidCode:=qidCode, _
                mfgLotNo:=mLot, _
                mfgLotNo2:=mLot2, _
                averyLabelCode:=averyLabelCode, _
                labelsPerSheet:=CStr(labelsPerSheet), _
                productID:=productID, _
                cpmuRef:=pmiNumber, _
                productDescription:=productDescription, _
                printType:="Reprint", _
                reprintReason:=reason, _
                originalPrintID:=printID, _
                reprintReasonType:=IIf(reasonMatch, "Standard", "User Entered Reason")) Then

                MsgBox "Warning: Failed to write reprint record to Label_Prints.csv", vbExclamation, "PRPrint"
            End If
        End If
    Next printRow

    Application.EnableEvents = True
    wsPrints.Range("C5").value = wsPrints.Range("C5").value
    
Cleanup:
    ExportNewLabelPrints
    AppStateUtils.RestoreAppState
    GoTo CleanExit

CleanupConnection:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    End If
    GoTo Cleanup

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PRPrint"
    GoTo CleanupConnection

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then Set rs = Nothing
    If Not conn Is Nothing Then Set conn = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    Set rngSub = Nothing
    Set rng = Nothing
    Set cell2Copy = Nothing
    Set cell1Copy = Nothing
    Set cleanTableRange = Nothing
    Set origTable = Nothing
    Set lblTable = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set reprintReason = Nothing
    Set labelPrintRow = Nothing
    Set printRow = Nothing
    Set tblReprintReasons = Nothing
    Set tblLabelPrintsImp = Nothing
    Set tblPrints = Nothing
    Set wsReprintReasons = Nothing
    Set wsLabelPrints = Nothing
    Set wsPrints = Nothing
    Exit Sub
End Sub
