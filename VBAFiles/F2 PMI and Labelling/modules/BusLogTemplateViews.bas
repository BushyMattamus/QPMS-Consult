Attribute VB_Name = "BusLogTemplateViews"
Sub BETempRev()
    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim pmiNumber As String
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim hWordWnd As LongPtr

    AppStateUtils.DisableAppUpdates

    ' --- Retrieve PMI Number from form ---
    pmiNumber = GetFieldValByLabelC("PMI Number:")
    If pmiNumber = "" Then
        MsgBox "Error: PMI Number cannot be empty.", vbExclamation, "BETempRev"
        GoTo CleanExit
    End If

    ' --- Construct folder path ---
    folderPath = GetFolderDrivePath(PRODUCT_APPROVED_LABELS_PATH)

    ' --- Verify folder exists ---
    If Not FolderExists(folderPath) Then
        MsgBox "Approved Labels folder path does not exist: " & folderPath, vbExclamation, "BETempRev"
        GoTo CleanExit
    End If

    ' --- Validate only one label file exists for this PMI Number ---
    If Not ValidateUniqueTemplateFile(folderPath, pmiNumber & "-", pmiNumber & "-*.docx", "Label", fileName) Then
        GoTo CleanExit
    End If

    fullPath = folderPath & "\" & fileName

    ' --- Open or get Word application ---
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo ErrHandler

    ' --- Configure Word visibility ---
    wdApp.Visible = True

    ' --- Bring Word window to front ---
    On Error Resume Next
    wdApp.Activate
    DoEvents
    On Error GoTo ErrHandler

    ' --- Open document ---
    If PROD_DEV = "DEV" Then
        Set wdDoc = wdApp.Documents.Open(fullPath)
    Else
        Set wdDoc = wdApp.Documents.Open(fullPath, ReadOnly:=True, AddToRecentFiles:=False)
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "BETempRev"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    AppStateUtils.RestoreAppState
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub
End Sub

Sub ViewPMI()
    On Error GoTo ErrHandler

    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    pmiNumber = GetFieldValByLabelH("PMI Number:")
    pmiVersion = GetFieldValByLabelH("PMI Version:")

    If Trim$(pmiNumber) = "" Then
        MsgBox "PMI Number cannot be empty.", vbExclamation, "ViewPMI"
        GoTo CleanExit
    End If

    If Trim$(pmiVersion) = "" Then
        MsgBox "PMI Version cannot be empty.", vbExclamation, "ViewPMI"
        GoTo CleanExit
    End If

    folderPath = GetFolderDrivePath(APPROVED_PMI_PDF_PATH)
    
    If Not FolderExists(folderPath) Then
        MsgBox "Approved PMI PDF folder path does not exist: " & folderPath, vbExclamation, "ViewPMI"
        GoTo CleanExit
    End If
    
    If Not ValidateUniqueTemplateFile(folderPath, pmiNumber, pmiNumber & "-" & pmiVersion & "_*.pdf", "PMI PDF", fileName) Then
        GoTo CleanExit
    End If
    
    fullPath = folderPath & fileName

    Shell "cmd /c start """" """ & fullPath & """", vbHide

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ViewPMI"
    GoTo CleanExit

CleanExit:
    Exit Sub
End Sub
