Attribute VB_Name = "BusLogPrintPMI"
' ========================================================================
' PMI AND LABEL PRINTING
' ========================================================================
' Handles printing of PMI documents with accompanying label
' Validates document control, prints PDF and label, archives copies
' ========================================================================

Option Explicit

' ========================================================================
' PMI PRINT MODE ENUMERATION
' ========================================================================
' Defines the available print modes for PMI document printing
' ========================================================================

Public Enum PrintMode
    pmLabelOnly = 1
    pmPdfOnly = 2
    pmBoth = 3
End Enum

' Module-level flag to track print completion
Private m_PrintCompleted As Boolean

' ========================================================================
' MAIN PMI PRINTING FUNCTION
' ========================================================================
' Prints PMI PDF and label together with validation and archiving
' ========================================================================

Public Sub PrintPMIDocuments()
    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim batchNumber As String
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim expiryDate As String
    Dim commentsValue As String
    Dim labelQuantity As Long
    Dim userName As String
    Dim wordFolderPath As String
    Dim pdfFolderPath As String
    Dim wordFileName As String
    Dim pdfFilename As String
    Dim wordFullPath As String
    Dim pdfFullPath As String
    Dim batchPreviouslyPrinted As Boolean
    Dim result As BatchStatusResult
    Dim updateComments As String
    Dim pdfPrintSuccess As Boolean
    Dim labelPrintSuccess As Boolean
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim response As VbMsgBoxResult
    Dim wsOriginal As Worksheet

    Set wsOriginal = ActiveSheet

    ' --- Retrieve input values from form ---
    batchNumber = GetFieldValByLabelH("Batch Number:")
    pmiNumber = GetFieldValByLabelH("PMI Number:")
    pmiVersion = GetFieldValByLabelH("PMI Version:")
    expiryDate = GetFieldValByLabelH("Expiry Date:")
    commentsValue = GetFieldValByLabelH("Comments:")
    userName = Environ("USERNAME")
    
    ' --- Validate no commas ---
    If Not ValidateNoCommas(commentsValue, "Comments") Then GoTo CleanExit
    
    ' --- Convert expiry date from dd/mm/yyyy to DD/MM/YY ---
    If IsDate(expiryDate) Then
        expiryDate = Format(CDate(expiryDate), "DD/MM/YY")
    End If

    ' --- Validate required fields ---
    If Trim$(batchNumber) = "" Then
        MsgBox "Batch Number cannot be empty.", vbExclamation, "PrintPMIDocuments"
        GoTo CleanExit
    End If

    If Trim$(pmiNumber) = "" Then
        MsgBox "PMI Number cannot be empty.", vbExclamation, "PrintPMIDocuments"
        GoTo CleanExit
    End If

    If Trim$(pmiVersion) = "" Then
        MsgBox "PMI Version cannot be empty.", vbExclamation, "PrintPMIDocuments"
        GoTo CleanExit
    End If

    If Trim$(expiryDate) = "" Then
        MsgBox "Expiry Date cannot be empty.", vbExclamation, "PrintPMIDocuments"
        GoTo CleanExit
    End If

    ' --- Check if batch was previously printed ---
    batchPreviouslyPrinted = HasBatchBeenPrinted(batchNumber)

    If batchPreviouslyPrinted Then
        response = MsgBox("This batch has already been printed. Click Yes to reprint, or No to cancel. If you select Yes, ensure to add a comment to allow printing.", _
                         vbYesNo + vbExclamation, "PrintPMIDocuments")
        If response = vbNo Then
            GoTo CleanExit
        End If

        If Trim$(commentsValue) = "" Then
            MsgBox "A comment is required when reprinting. Please add a comment explaining the reason for reprinting.", _
                   vbExclamation, "PrintPMIDocuments"
            GoTo CleanExit
        End If

        result = ValidateAndHandleBatchStatusReprint(batchNumber, "PrintPMIDocuments")
    Else
        result = ValidateAndHandleBatchStatusNew(batchNumber, "PrintPMIDocuments")
    End If

    If result <> bsrAllow Then
        GoTo CleanExit
    End If

    ' --- Construct folder paths ---
    wordFolderPath = GetFolderDrivePath(APPROVED_PMI_WORD_PATH)
    pdfFolderPath = GetFolderDrivePath(APPROVED_PMI_PDF_PATH)

    ' --- Verify Word folder exists ---
    If Not FolderExists(wordFolderPath) Then
        MsgBox "Approved PMI Word folder path does not exist: " & wordFolderPath, vbExclamation, "PrintPMIDocuments"
        GoTo CleanExit
    End If

    ' --- Verify PDF folder exists ---
    If Not FolderExists(pdfFolderPath) Then
        MsgBox "Approved PMI PDF folder path does not exist: " & pdfFolderPath, vbExclamation, "PrintPMIDocuments"
        GoTo CleanExit
    End If

    ' --- Validate unique Word template file ---
    If Not ValidateUniqueTemplateFile(wordFolderPath, pmiNumber, pmiNumber & "-" & pmiVersion & "_*.docx", "PMI Word", wordFileName) Then
        GoTo CleanExit
    End If

    wordFullPath = wordFolderPath & wordFileName

    ' --- Validate unique PDF file ---
    If Not ValidateUniqueTemplateFile(pdfFolderPath, pmiNumber, pmiNumber & "-" & pmiVersion & "*.pdf", "PMI PDF", pdfFilename) Then
        GoTo CleanExit
    End If

    pdfFullPath = pdfFolderPath & pdfFilename

    ' --- Open Word document for version validation ---
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo ErrHandler

    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(wordFullPath, ReadOnly:=True)

    ' --- Validate PMI Version in document header ---
    If Not ValidatePMIDocumentVersion(wdDoc, wdApp, pmiVersion) Then
        wdDoc.Close SaveChanges:=False
        If wdApp.Documents.Count = 0 Then wdApp.Quit
        GoTo CleanExit
    End If

    ' --- Extract page count for label quantity ---
    labelQuantity = wdDoc.ComputeStatistics(2)
    
    If labelQuantity <= 0 Then
        MsgBox "Unable to determine page count from PMI document.", vbExclamation, "PrintPMIDocuments"
        wdDoc.Close SaveChanges:=False
        If wdApp.Documents.Count = 0 Then wdApp.Quit
        GoTo CleanExit
    End If

    ' --- Close Word document after validation ---
    wdDoc.Close SaveChanges:=False
    If wdApp.Documents.Count = 0 Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    ' --- Initialize print success flags ---
    pdfPrintSuccess = False
    labelPrintSuccess = False

    ' --- Execute prints based on mode ---
    pdfPrintSuccess = PrintPMIPdf(pdfFullPath, pmiNumber, pmiVersion, batchNumber)
    If pdfPrintSuccess Then
        labelPrintSuccess = PrintPMILabel(batchNumber, pmiNumber, pmiVersion, expiryDate, userName, labelQuantity)
    End If

    wsOriginal.Activate

    updateComments = BuildPrintComments(batchPreviouslyPrinted, pdfPrintSuccess, labelPrintSuccess, commentsValue)

    ' --- Update batch status if both prints succeeded ---
    If pdfPrintSuccess And labelPrintSuccess Then
        UpdateFinishedGoodsBatchRecord batchNumber, batchStatus:=FGB_STATUS_PRINTED, _
            dateUpdated:=Format(Date, "YYYY-MM-DD"), updatedBy:=userName, updateComments:=updateComments
        MsgBox "PMI documents printed successfully for batch: " & batchNumber, vbInformation, "PrintPMIDocuments"
    Else
        UpdateFinishedGoodsBatchRecord batchNumber, batchStatus:="", _
            dateUpdated:=Format(Date, "YYYY-MM-DD"), updatedBy:=userName, updateComments:=updateComments
        MsgBox "Print operation incomplete for batch: " & batchNumber & vbCrLf & vbCrLf & updateComments, vbExclamation, "PrintPMIDocuments"
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PrintPMIDocuments"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not wdDoc Is Nothing Then
        wdDoc.Close SaveChanges:=False
        Set wdDoc = Nothing
    End If
    If Not wdApp Is Nothing Then
        If wdApp.Documents.Count = 0 Then wdApp.Quit
        Set wdApp = Nothing
    End If
    Set wsOriginal = Nothing
End Sub

' ========================================================================
' PRINT PMI PDF
' ========================================================================
' Prints the approved PMI PDF to the default printer and archives a copy
' Returns True if print was initiated successfully
' ========================================================================

Private Function PrintPMIPdf(ByVal pdfFullPath As String, ByVal pmiNumber As String, _
                             ByVal pmiVersion As String, ByVal batchNumber As String) As Boolean
    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    ' Dim archiveFolderPath As String
    ' Dim archiveFileName As String
    ' Dim archiveFullPath As String
    ' Dim fso As Object

    PrintPMIPdf = False

    ' --- Verify source PDF exists ---
    If Dir(pdfFullPath) = "" Then
        MsgBox "PDF file not found: " & pdfFullPath, vbExclamation, "PrintPMIPdf"
        GoTo CleanExit
    End If

    ' --- Construct archive folder path ---
    ' archiveFolderPath = GetFolderDrivePath(PRODUCTION_ARCHIVE_PMI_PATH)

    ' If Dir(archiveFolderPath, vbDirectory) = "" Then
    '     MsgBox "Archive PMI Labels folder path does not exist: " & archiveFolderPath, vbExclamation, "PrintPMIPdf"
    '     GoTo CleanExit
    ' End If

    ' --- Build archive filename ---
    ' archiveFileName = batchNumber & "_" & pmiNumber & "-" & pmiVersion & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
    ' archiveFullPath = archiveFolderPath & archiveFileName

    ' --- Copy PDF to archive ---
    ' Set fso = CreateObject("Scripting.FileSystemObject")
    ' fso.CopyFile pdfFullPath, archiveFullPath, True

    ' --- Open PDF in default viewer ---
    CreateObject("Shell.Application").ShellExecute pdfFullPath

    PrintPMIPdf = True

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PrintPMIPdf"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    ' Set fso = Nothing
End Function

' ========================================================================
' PRINT PMI LABEL
' ========================================================================
' Populates the PMI_Label worksheet and prints with user printer selection
' Returns True if print was completed, False if cancelled
' ========================================================================

Private Function PrintPMILabel(ByVal batchNumber As String, ByVal pmiNumber As String, _
                               ByVal pmiVersion As String, ByVal expiryDate As String, _
                               ByVal userName As String, ByVal labelQuantity As Long) As Boolean
    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim wsPMILabel As Worksheet
    Dim archiveFolderPath As String
    Dim archiveFileName As String
    Dim archiveFullPath As String
    Dim barcodeValue As String
    Dim currentLabel As Long
    Dim dialogResult As Boolean
    Dim originalVisibility As XlSheetVisibility

    PrintPMILabel = False

    ' --- Get PMI_Label worksheet ---
    On Error Resume Next
    Set wsPMILabel = ThisWorkbook.Worksheets("PMI_Label")
    On Error GoTo ErrHandler

    If wsPMILabel Is Nothing Then
        MsgBox "PMI_Label worksheet not found.", vbExclamation, "PrintPMILabel"
        GoTo CleanExit
    End If

    ' --- Store original visibility and unhide if necessary ---
    originalVisibility = wsPMILabel.Visible
    If wsPMILabel.Visible <> xlSheetVisible Then
        wsPMILabel.Visible = xlSheetVisible
    End If

    ' --- Build barcode value with asterisks ---
    barcodeValue = "*" & batchNumber & "*"

    wsPMILabel.Activate

    wsPMILabel.Range("A1").value = barcodeValue
    wsPMILabel.Range("A2").value = "Bx: " & batchNumber
    wsPMILabel.Range("A3").value = "Exp.: " & expiryDate
    wsPMILabel.Range("A4").value = "User: " & userName
    wsPMILabel.Range("A5").value = "PMI: " & pmiNumber & "-" & pmiVersion & " 1/" & labelQuantity

    ' --- Let user select printer ---
    dialogResult = Application.Dialogs(xlDialogPrinterSetup).Show

    If Not dialogResult Then
        MsgBox "Printer selection was cancelled by user.", vbInformation, "PrintPMILabel"
        GoTo CleanExit
    End If

    ' --- Print first label ---
    On Error Resume Next
    wsPMILabel.PrintOut Copies:=1
    If Err.Number <> 0 Then
        MsgBox "Error printing first label: " & Err.Description, vbExclamation, "PrintPMILabel"
        Err.Clear
        GoTo CleanExit
    End If
    On Error GoTo ErrHandler

    ' --- Print remaining labels using same printer settings ---
    On Error Resume Next
    For currentLabel = 2 To labelQuantity
        wsPMILabel.Range("A5").value = "PMI: " & pmiNumber & "-" & pmiVersion & " " & currentLabel & "/" & labelQuantity
        wsPMILabel.PrintOut Copies:=1
        If Err.Number <> 0 Then
            MsgBox "Error printing label " & currentLabel & " of " & labelQuantity & ": " & Err.Description, vbExclamation, "PrintPMILabel"
            Err.Clear
            Exit For
        End If
        DoEvents
    Next currentLabel
    On Error GoTo ErrHandler

    ' --- Construct archive folder path ---
    archiveFolderPath = GetFolderDrivePath(ARCHIVE_PMI_LABELS_PATH)

    If Dir(archiveFolderPath, vbDirectory) = "" Then
        MsgBox "Archive Labels folder path does not exist: " & archiveFolderPath, vbExclamation, "PrintPMILabel"
        GoTo CleanExit
    End If

    ' --- Build archive filename ---
    archiveFileName = "PMILabel_" & batchNumber & "_" & pmiNumber & "-" & pmiVersion & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
    archiveFullPath = archiveFolderPath & archiveFileName

    wsPMILabel.Range("A5").value = "PMI: " & pmiNumber & "-" & pmiVersion & " 1/" & labelQuantity

    wsPMILabel.ExportAsFixedFormat Type:=xlTypePDF, fileName:=archiveFullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    PrintPMILabel = True

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PrintPMILabel"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not wsPMILabel Is Nothing Then
        wsPMILabel.Visible = originalVisibility
    End If
    Set wsPMILabel = Nothing
End Function

' ========================================================================
' BUILD PRINT COMMENTS
' ========================================================================
' Constructs the appropriate comment string based on print results
' ========================================================================

Private Function BuildPrintComments(ByVal isReprint As Boolean, ByVal pdfSuccess As Boolean, _
                                    ByVal labelSuccess As Boolean, ByVal userComments As String) As String
    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim baseComment As String
    Dim statusText As String

    If pdfSuccess And labelSuccess Then
        If isReprint Then
            statusText = "Reprint PMI Print (Label + PDF)"
        Else
            statusText = "Original PMI Print (Label + PDF)"
        End If
    Else
        statusText = "Incomplete PMI Print (Label + PDF)"
        If Not pdfSuccess Then
            statusText = statusText & " - PDF failed"
        End If
        If Not labelSuccess Then
            statusText = statusText & " - Label cancelled"
        End If
    End If

    If Trim$(userComments) <> "" Then
        BuildPrintComments = statusText & " - " & userComments
    Else
        BuildPrintComments = statusText
    End If

    GoTo CleanExit

ErrHandler:
    BuildPrintComments = "Error building comments"
    GoTo CleanExit

CleanExit:
End Function
