Attribute VB_Name = "BusLogRecordLoader"
' filepath: c:\Users\CollettMatt\OneDrive - Queensland Health\Interim State Uplift\03. Delivery\31. VBA Development\VBA Codes\CPISU 14 RM\F2 PMI and Label Printing\BusLogRecordLoader.bas
Option Explicit

Public Sub LoadBatchDetailsFromMaster()
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    Dim batchNumber As String
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim productDescription As String
    Dim expiryDate As String

    batchNumber = Trim$(GetFieldValByLabelH("Batch Number:"))
    
    If batchNumber = "" Then
        MsgBox "Please enter a Batch Number.", vbExclamation, "Missing Batch Number"
        GoTo CleanExit
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPath = CreateTemporaryFileCopy(FILE_FINISHED_GOODS_BATCH_MASTER)
    
    If tempPath = "" Then
        MsgBox "Failed to create temporary copy of batch master file.", vbCritical, "LoadBatchDetailsFromMaster"
        GoTo CleanExit
    End If
    
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    Set conn = OpenADOConnection(folderPath)

    If conn Is Nothing Then
        MsgBox "Failed to open connection to temporary file.", vbCritical, "LoadBatchDetailsFromMaster"
        GoTo CleanExit
    End If

    sql = "SELECT [" & FGBNM_HEADER_PMI_NUMBER & "], " & _
          "[" & FGBNM_HEADER_PMI_VERSION & "], " & _
          "[" & FGBNM_HEADER_PRODUCT_DESCRIPTION & "], " & _
          "[" & FGBNM_HEADER_EXPIRY_DATE & "] " & _
          "FROM [" & fileName & "] " & _
          "WHERE [" & FGBNM_HEADER_BATCH_NUMBER & "] = '" & Replace(batchNumber, "'", "''") & "'"

    Set rs = conn.Execute(sql)

    If rs.EOF Then
        MsgBox "Batch Number " & batchNumber & " not found.", vbExclamation, "Batch Not Found"
        Range("I6:I9").ClearContents
        GoTo CleanExit
    End If

    pmiNumber = Trim$(rs.Fields(FGBNM_HEADER_PMI_NUMBER).value & "")
    pmiVersion = Trim$(rs.Fields(FGBNM_HEADER_PMI_VERSION).value & "")
    productDescription = Trim$(rs.Fields(FGBNM_HEADER_PRODUCT_DESCRIPTION).value & "")
    expiryDate = Trim$(rs.Fields(FGBNM_HEADER_EXPIRY_DATE).value & "")

    Range("I5").value = productDescription
    Range("I7").value = pmiNumber
    Range("I8").value = pmiVersion
    Range("I9").value = CStr(expiryDate)
    

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LoadBatchDetailsFromMaster"
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Sub
End Sub

Public Sub LookupPMINumberByProductDescription(ByVal productDescription As String)
    On Error GoTo ErrHandler

    Dim wsFGMaster As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim foundPMI As String
    Dim matchCount As Long

    If Trim$(productDescription) = "" Then
        GoTo CleanExit
    End If

    Set wsFGMaster = ThisWorkbook.Worksheets("FG_Master")

    lastRow = wsFGMaster.Cells(wsFGMaster.Rows.Count, "F").End(xlUp).row

    If lastRow < 2 Then
        MsgBox "No data found in FG_Master worksheet.", vbExclamation, "LookupPMINumberByProductDescription"
        GoTo CleanExit
    End If

    matchCount = 0
    foundPMI = ""

    For i = 2 To lastRow
        If StrComp(Trim$(wsFGMaster.Cells(i, "F").value), productDescription, vbTextCompare) = 0 Then
            If StrComp(Trim$(wsFGMaster.Cells(i, "G").value), "Current", vbTextCompare) = 0 Then
                matchCount = matchCount + 1
                If matchCount = 1 Then
                    foundPMI = Trim$(wsFGMaster.Cells(i, "A").value)
                End If
            End If
        End If
    Next i

    If matchCount = 0 Then
        MsgBox "No current PMI found for product: " & productDescription, vbInformation, "LookupPMINumberByProductDescription"
        Range("D6").ClearContents
        GoTo CleanExit
    End If

    If matchCount > 1 Then
        MsgBox "Multiple current PMI records found for product: " & productDescription & vbCrLf & _
               "Please select the correct PMI manually.", vbExclamation, "LookupPMINumberByProductDescription"
        Range("D6").ClearContents
        GoTo CleanExit
    End If

    Range("D6").value = foundPMI

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LookupPMINumberByProductDescription"
    GoTo CleanExit

CleanExit:
    Set wsFGMaster = Nothing
    Exit Sub
End Sub

Public Sub LookupProductDescriptionByPMINumber(ByVal pmiNumber As String)
    On Error GoTo ErrHandler

    Dim wsFGMaster As Worksheet
    Dim wsLabelReg As Worksheet
    Dim tblLabelReg As ListObject
    Dim lastRow As Long
    Dim i As Long
    Dim foundDescription As String
    Dim matchCount As Long
    Dim foundRow As ListRow
    Dim labelsPerBatch As String
    Dim specialInstructions As String

    If Trim$(pmiNumber) = "" Then
        GoTo CleanExit
    End If

    Set wsFGMaster = ThisWorkbook.Worksheets("FG_Master")

    lastRow = wsFGMaster.Cells(wsFGMaster.Rows.Count, "A").End(xlUp).row

    If lastRow < 2 Then
        MsgBox "No data found in FG_Master worksheet.", vbExclamation, "LookupProductDescriptionByPMINumber"
        GoTo CleanExit
    End If

    matchCount = 0
    foundDescription = ""

    For i = 2 To lastRow
        If StrComp(Trim$(wsFGMaster.Cells(i, "A").value), pmiNumber, vbTextCompare) = 0 Then
            If StrComp(Trim$(wsFGMaster.Cells(i, "G").value), "Current", vbTextCompare) = 0 Then
                matchCount = matchCount + 1
                If matchCount = 1 Then
                    foundDescription = Trim$(wsFGMaster.Cells(i, "F").value)
                End If
            End If
        End If
    Next i

    If matchCount = 0 Then
        MsgBox "No current product found for PMI: " & pmiNumber, vbInformation, "LookupProductDescriptionByPMINumber"
        Range("D5").ClearContents
        GoTo CleanExit
    End If

    If matchCount > 1 Then
        MsgBox "Multiple current product records found for PMI: " & pmiNumber & vbCrLf & _
               "Please select the correct product manually.", vbExclamation, "LookupProductDescriptionByPMINumber"
        Range("D5").ClearContents
        GoTo CleanExit
    End If

    Range("D5").value = foundDescription

    Set wsLabelReg = ThisWorkbook.Worksheets("Label Register")
    Set tblLabelReg = wsLabelReg.ListObjects("tblLabelReg")

    labelsPerBatch = ""
    specialInstructions = ""

    For Each foundRow In tblLabelReg.ListRows
        If StrComp(Trim$(foundRow.Range(1, 1).value), pmiNumber, vbTextCompare) = 0 Then
            labelsPerBatch = Trim$(foundRow.Range(1, 6).value)
            specialInstructions = Trim$(foundRow.Range(1, 5).value)
            Exit For
        End If
    Next foundRow

    If labelsPerBatch <> "" Then
        Range("D7").value = labelsPerBatch
    Else
        Range("D7").ClearContents
    End If

    If specialInstructions <> "" Then
        Range("D9").value = specialInstructions
    Else
        Range("D9").ClearContents
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "LookupProductDescriptionByPMINumber"
    GoTo CleanExit

CleanExit:
    Set foundRow = Nothing
    Set tblLabelReg = Nothing
    Set wsLabelReg = Nothing
    Set wsFGMaster = Nothing
    Exit Sub
End Sub
