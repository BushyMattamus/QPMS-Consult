Attribute VB_Name = "BusLogFileViewing"
' ========================================================================
' GENERIC FILE VIEWING
' ========================================================================
' Opens archived files based on selection in a table
' Searches specified folder for matching files using reference column
' ========================================================================

Option Explicit

Sub ViewSelectedPDFs(ByVal folderPath As String, _
                     ByVal worksheetName As String, _
                     ByVal tableName As String, _
                     ByVal referenceColumn As String, _
                     ByVal viewColumn As String)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim referenceID As String
    Dim archiveFolder As String
    Dim fileName As String
    Dim foundFiles As Collection
    Dim filePath As Variant
    Dim selectedCount As Integer
    Dim viewedCount As Integer
    
    Set foundFiles = New Collection
    
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    Set tbl = ws.ListObjects(tableName)
    
    archiveFolder = GetFolderDrivePath(folderPath)
    
    If Dir(archiveFolder, vbDirectory) = "" Then
        MsgBox "Archive folder not found: " & archiveFolder, vbExclamation, "View Files"
        GoTo CleanExit
    End If
    
    selectedCount = 0
    viewedCount = 0
    
    For Each row In tbl.ListRows
        If LCase(row.Range.Cells(1, tbl.ListColumns(viewColumn).Index).value) = "x" Then
            selectedCount = selectedCount + 1
            
            referenceID = row.Range.Cells(1, tbl.ListColumns(referenceColumn).Index).value
            
            If Trim(referenceID) <> "" Then
                fileName = Dir(archiveFolder & "*" & referenceID & "*.pdf")
                
                Do While fileName <> ""
                    foundFiles.Add archiveFolder & fileName
                    viewedCount = viewedCount + 1
                    fileName = Dir()
                Loop
                
                If foundFiles.Count = 0 Then
                    MsgBox "PDF file not found for " & referenceColumn & ": " & referenceID, vbExclamation, "View Files"
                End If
            End If
        End If
    Next row
    
    If selectedCount = 0 Then
        MsgBox "No records selected. Please check the '" & viewColumn & "' column for records you want to view.", vbInformation, "View Files"
        GoTo CleanExit
    End If
    
    If foundFiles.Count = 0 Then
        MsgBox "No PDF files found for the selected records.", vbExclamation, "View Files"
        GoTo CleanExit
    End If
    
    For Each filePath In foundFiles
        OpenPDFFile CStr(filePath)
    Next filePath
    
    If viewedCount > 0 Then
        MsgBox viewedCount & " PDF file(s) opened successfully.", vbInformation, "View Files"
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ViewSelectedPDFs"
    GoTo CleanExit

CleanExit:
    Set foundFiles = Nothing
    Set row = Nothing
    Set tbl = Nothing
    Set ws = Nothing
    Exit Sub
End Sub

Private Sub OpenPDFFile(ByVal filePath As String)
    On Error GoTo ErrHandler
    
    Shell "explorer.exe """ & filePath & """", vbNormalFocus
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error opening PDF file: " & filePath & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "OpenPDFFile"
    GoTo CleanExit

CleanExit:
    Exit Sub
End Sub
