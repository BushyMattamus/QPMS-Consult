Attribute VB_Name = "CreateBatch"
Option Explicit


Public Sub WriteBatchToCSV(sheetName As String)
    On Error GoTo ErrHandler

    ' --- Variable declarations ---
    Dim ws As Worksheet
    Dim masterCSV As String
    Dim auditCSV As String
    Dim Batch_Number As String
    Dim PMI_Number As String
    Dim PMI_Version As String
    Dim Product_Description As String
    Dim Shelf_No As String
    Dim Shelf_Life As String
    Dim Manufacturing_Date As String
    Dim Expiry_Date As String
    Dim Date_Created As String
    Dim Time_Created As String
    Dim Batch_Status As String
    Dim Created_By As String
    Dim Date_Updated As String
    Dim Updated_By As String
    Dim Update_Comments As String
    Dim scanned As String
    Dim missing As String
    Dim Trans_Id As String
    Dim masterRow As String
    Dim auditRow As String
    Dim userResponse As VbMsgBoxResult

    Set ws = ThisWorkbook.Sheets(sheetName)

    ' --- File paths ---
    masterCSV = GetDataFilePath(FOLDER_MATERIAL_MANAGEMENT & "\" & FOLDER_TABLES, FILE_FINISHED_GOODS_BATCH_NUMBER_MASTER)
    auditCSV = GetDataFilePath(FOLDER_MATERIAL_MANAGEMENT & "\" & FOLDER_TABLES, FILE_FINISHED_GOODS_BATCH_NUMBER_AUDIT)

    ' --- Pull all values from the Batch Number Generation sheet ---
    Batch_Number = GetFieldByLabel(sheetName, "C", "Batch Number").value
    PMI_Number = GetFieldByLabel(sheetName, "C", "PMI number").value
    PMI_Version = GetFieldByLabel(sheetName, "C", "PMI version").value
    Product_Description = GetFieldByLabel(sheetName, "C", "Product Description").value
    Shelf_No = GetFieldByLabel(sheetName, "C", "Shelf life(No of days/Months/Years)").value
    Shelf_Life = GetFieldByLabel(sheetName, "C", "Shelf life(Days/Months/Years)").value
    Manufacturing_Date = GetFieldByLabel(sheetName, "C", "Manufacturing date").value
    Expiry_Date = GetFieldByLabel(sheetName, "C", "Recommended expiry date").value

    Date_Created = Format(Date, "dd/mm/yyyy")
    Time_Created = Format(Time, "HH:MM:SS")
    Batch_Status = "Created"
    Created_By = Environ$("Username")

    Date_Updated = ""
    Updated_By = ""
    Update_Comments = ""

    ' --- Validation block ---
    If Trim(Batch_Number) = "" Then missing = missing & vbCrLf & "- Batch Number"
    If Trim(PMI_Number) = "" Then missing = missing & vbCrLf & "- PMI Number"
    If Trim(PMI_Version) = "" Then missing = missing & vbCrLf & "- PMI Version"
    If Trim(Product_Description) = "" Then missing = missing & vbCrLf & "- Product Description"

    If UCase(Trim(Shelf_Life)) <> "AS PER MANUFACTURER" Then
        If Trim(Shelf_No) = "" Then missing = missing & vbCrLf & "- Shelf Life (Number)"
    End If

    If Trim(Shelf_Life) = "" Then missing = missing & vbCrLf & "- Shelf Life (Type)"
    If Trim(Manufacturing_Date) = "" Then missing = missing & vbCrLf & "- Manufacturing Date"
    If Trim(Expiry_Date) = "" Then missing = missing & vbCrLf & "- Recommended Expiry Date"

    If missing <> "" Then
        MsgBox "The following required fields are missing:" & vbCrLf & missing, vbCritical, "Missing Data"
        GoTo CleanExit
    End If
    
    ' --- Comma validation for all fields being written ---
    If Not ValidateNoCommas(Batch_Number, "Batch Number") Then GoTo CleanExit
    If Not ValidateNoCommas(PMI_Number, "PMI Number") Then GoTo CleanExit
    If Not ValidateNoCommas(PMI_Version, "PMI Version") Then GoTo CleanExit
    If Not ValidateNoCommas(Product_Description, "Product Description") Then GoTo CleanExit
    If Not ValidateNoCommas(Shelf_No, "Shelf life(No of days/Months/Years)") Then GoTo CleanExit
    If Not ValidateNoCommas(Shelf_Life, "Shelf life(Days/Months/Years)") Then GoTo CleanExit

    ' --- Validate expiry date calculation ---
    If Not ValidateExpiryDate(Manufacturing_Date, Expiry_Date, Shelf_No, Shelf_Life, userResponse) Then
        If userResponse = vbNo Then
            GoTo CleanExit
        End If
    End If

    ' --- Convert dates to ANSI format ---
    Manufacturing_Date = DateToANSI(Manufacturing_Date)
    Expiry_Date = DateToANSI(Expiry_Date)
    Date_Created = DateToANSI(Date_Created)

    ' --- Get transaction ID for audit ---
    Trans_Id = GetNextTransID(auditCSV)

    ' --- Construct CSV rows ---
    masterRow = """" & Replace(Batch_Number, """", """""") & """," & _
                """" & Replace(PMI_Number, """", """""") & """," & _
                """" & Replace(PMI_Version, """", """""") & """," & _
                """" & Replace(Product_Description, """", """""") & """," & _
                """" & Replace(Shelf_No, """", """""") & """," & _
                """" & Replace(Shelf_Life, """", """""") & """," & _
                """" & Replace(Manufacturing_Date, """", """""") & """," & _
                """" & Replace(Expiry_Date, """", """""") & """," & _
                """" & Replace(Date_Created, """", """""") & """," & _
                """" & Replace(Time_Created, """", """""") & """," & _
                """" & Replace(Batch_Status, """", """""") & """," & _
                """" & Replace(Created_By, """", """""") & """," & _
                """" & Replace(Date_Updated, """", """""") & """," & _
                """" & Replace(Updated_By, """", """""") & """," & _
                """" & Replace(Update_Comments, """", """""") & """," & _
                """" & Replace(scanned, """", """""") & """"

    auditRow = masterRow & "," & """" & Replace(Trans_Id, """", """""") & """"

    ' --- Verify batch number uniqueness ---
    If Not IsBatchNumberUnique(Batch_Number, masterCSV) Then
        MsgBox "Batch Number '" & Batch_Number & "' already exists in the master file." & vbCrLf & _
               "A new batch number will be generated for you, but you have to resubmit the batch.", vbCritical, "Duplicate Batch Number"
        GoTo CreateNextBatchNumber
    End If

    ' --- Append to CSV files ---
    AppendLineToFile masterCSV, masterRow
    AppendLineToFile auditCSV, auditRow

    ' --- Refresh data and generate next batch number ---
    Call ImportSelectedCSVColumns
    Call ImportFgBatchSelectedCSVColumns
    ClearFields

    ' Next Batch Number
    GetFieldByLabel("Batch Number Generation", "C", "Batch number").value = GenerateBatchNumber
    GetFieldByLabel(sheetName, "C", "Product description").ClearContents

    MsgBox "Batch created successfully.", vbInformation
    GoTo CleanExit

CreateNextBatchNumber:
    ' --- Placeholder for future functionality ---
    GetFieldByLabel("Batch Number Generation", "C", "Batch number").value = GenerateBatchNumber
    GoTo CleanExit

ErrHandler:
    MsgBox "Error writing CSV files: " & Err.Description, vbCritical
    GoTo CleanExit

CleanExit:
    Set ws = Nothing
    Exit Sub
End Sub

' ========================================================================
' Verify that a batch number is unique in the master CSV file
' ========================================================================
Private Function IsBatchNumberUnique(ByVal batchNumber As String, ByVal csvFilePath As String) As Boolean
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim rs As Object
    Dim tempFilePath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim recordCount As Long
    Dim fso As Object

    IsBatchNumberUnique = False

    ' --- Create temporary copy for safe reading ---
    tempFilePath = CreateTemporaryFileCopy(FILE_FINISHED_GOODS_BATCH_NUMBER_MASTER)
    If tempFilePath = "" Then
        MsgBox "Unable to create temporary file for batch validation.", vbCritical, "IsBatchNumberUnique"
        GoTo CleanExit
    End If

    ' --- Extract folder and filename ---
    folderPath = Left(tempFilePath, InStrRev(tempFilePath, "\") - 1)
    fileName = Mid(tempFilePath, InStrRev(tempFilePath, "\") + 1)

    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(folderPath)
    If conn Is Nothing Then
        GoTo CleanExit
    End If

    ' --- Query for batch number ---
    Set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT COUNT(*) AS RecordCount FROM [" & fileName & "] WHERE " & FGBNM_HEADER_BATCH_NUMBER & " = '" & Replace(batchNumber, "'", "''") & "'"

    rs.Open sql, conn

    If Not rs.EOF Then
        recordCount = rs.fields("RecordCount").value
        IsBatchNumberUnique = (recordCount = 0)
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "IsBatchNumberUnique"
    IsBatchNumberUnique = False
    GoTo CleanExit

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If

    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If

    ' --- Clean up temporary file ---
    If tempFilePath <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(tempFilePath) Then
            On Error Resume Next
            fso.DeleteFile tempFilePath, True
            On Error GoTo 0
        End If
        Set fso = Nothing
    End If

    Exit Function
End Function
