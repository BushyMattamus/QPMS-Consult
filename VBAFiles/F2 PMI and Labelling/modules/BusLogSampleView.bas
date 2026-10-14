Attribute VB_Name = "BusLogSampleView"
Sub SampLabel()
    ' Sample Label View - Uses SavePrintLabel with sample view mode
    ' Displays 10 sample labels in Word for review without printing or saving
    
    SavePrintLabel isSampleView:=True
    
End Sub

Sub SampleViewPMI()
' Opens the approved PMI document and populates batch details in the header only.

    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim batchNumber As String
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim expiryDate As String
    Dim userName As String
    Dim barcodeValue As String
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wdSection As Object
    Dim wdHeader As Object
    Dim hWordWnd As LongPtr

    ' --- Retrieve input values from form ---
    batchNumber = GetFieldValByLabelH("Batch Number:")
    pmiNumber = GetFieldValByLabelH("PMI Number:")
    pmiVersion = GetFieldValByLabelH("PMI Version:")
    expiryDate = GetFieldValByLabelH("Expiry Date:")
    userName = Environ("USERNAME")

    ' --- Build barcode value with asterisks for BC C39 ---
    barcodeValue = "*" & batchNumber & "*"

    ' --- Construct file path ---
    folderPath = GetFolderDrivePath(APPROVED_PMI_PATH)
    
    ' --- Verify folder exists ---
    If Not FolderExists(folderPath) Then
        MsgBox "Approved PMI folder path does not exist: " & folderPath, vbExclamation, "SampleViewPMI"
        GoTo CleanExit
    End If

    ' --- Validate unique PMI file exists ---
    If Not ValidateUniqueTemplateFile(folderPath, pmiNumber, pmiNumber & "-" & pmiVersion & "_*.docx", "PMI", fileName) Then
        GoTo CleanExit
    End If
    
    fullPath = folderPath & fileName

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

    ' --- Open document as read-only ---
    Set wdDoc = wdApp.Documents.Open(fullPath, ReadOnly:=True)

    ' --- Validate PMI version in document header ---
    If Not ValidatePMIDocumentVersion(wdDoc, wdApp, pmiVersion) Then
        GoTo CleanExit
    End If

    ' --- Replace placeholders in headers only ---
    For Each wdSection In wdDoc.Sections
        For Each wdHeader In wdSection.headers
            With wdHeader.Range.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Forward = True
                .Wrap = 0 ' wdFindStop
                .MatchCase = True
                .MatchWholeWord = False

                ' Replace batch number placeholder
                .Text = TEXT_HOLDER_BATCH
                .Replacement.Text = batchNumber
                .Execute Replace:=2 ' wdReplaceAll

                ' Replace barcode placeholder
                .Text = TEXT_HOLDER_BARCODE
                .Replacement.Text = barcodeValue
                .Execute Replace:=2

                ' Replace expiry placeholder
                .Text = TEXT_HOLDER_EXP
                .Replacement.Text = expiryDate
                .Execute Replace:=2

                ' Append username after "Batch Issued by:"
                .Text = TEXT_HOLDER_ISSUED_BY
                .Replacement.Text = TEXT_HOLDER_ISSUED_BY & " " & userName
                .Execute Replace:=2
            End With
        Next wdHeader
    Next wdSection

    ' --- Update fields in document ---
    wdDoc.Fields.Update

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "SampleViewPMI"
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Set wdHeader = Nothing
    Set wdSection = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
