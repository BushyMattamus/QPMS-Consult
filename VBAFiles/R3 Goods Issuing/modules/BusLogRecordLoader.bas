Attribute VB_Name = "BusLogRecordLoader"
Option Explicit

Public Sub LoadQCNumberRecord()
    ' Get the QC number
    Dim qcNumber As String
    qcNumber = GetFieldValByLabelC("QC Number:")
    
    If qcNumber = "" Then
        MsgBox "Please enter a QC Number in cell D3", vbExclamation
        Exit Sub
    End If
    
    ' Get the file path
    Dim csvPath As String
    csvPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not FileExists(csvPath) Then
        MsgBox "QC Number database file not found: " & csvPath, vbExclamation
        Exit Sub
    End If
    
    ' Read and process the CSV file
    Dim fileNum As Integer
    Dim headerLine As String
    Dim dataLine As String
    Dim headers() As String
    Dim values() As String
    Dim foundRecord As Boolean
    
    fileNum = FreeFile
    Open csvPath For Input As #fileNum
    
    ' Read headers
    Line Input #fileNum, headerLine
    headers = Split(headerLine, ",")
    
    ' Search for matching QC number
    foundRecord = False
    Do Until EOF(fileNum)
        Line Input #fileNum, dataLine
        values = Split(dataLine, ",")
        
        ' Check if this is our QC number (first column)
        If Len(values(0)) > 0 Then
            If Trim$(Replace(values(0), """", "")) = qcNumber Then
                foundRecord = True
                Exit Do
            End If
        End If
    Loop
    
    Close #fileNum
    
    If Not foundRecord Then
        MsgBox "QC Number " & qcNumber & " not found in database.", vbExclamation
        Exit Sub
    End If
    
    ' Clear and Populate the form fields

    Range("D4:D25").ClearContents
    
    ' Map the values to their respective fields
    Dim i As Integer
    Dim lastRetestDate As String
    Dim firstRetestDate As String
    
    For i = 0 To UBound(headers)
        If i > UBound(values) Then Exit For   ' safety guard
        
        Dim cleanValue As String
        cleanValue = Replace(values(i), """", "")
        
        Select Case Trim$(headers(i))
            Case HEADER_STANDARD_NAME
                Range("D4").value = cleanValue
            Case HEADER_RM_CODE
                Range("D5").value = Format$(cleanValue, "0000")
            Case HEADER_BRAND_NAME
                Range("D6").value = cleanValue
            Case HEADER_QUICK_CODE
                Range("D7").value = cleanValue
            Case HEADER_PART_ARTG
                Range("D8").value = cleanValue
            Case HEADER_VENDOR_NAME
                Range("D9").value = cleanValue
            Case HEADER_MANUFACTURER_NAME
                Range("D10").value = cleanValue
            Case HEADER_MANUFACTURER_LOT
                Range("D11").value = cleanValue
            Case HEADER_MANUFACTURER_EXPIRY
                SetDateStringInCell Range("D12"), ANSIToDate(cleanValue)
            Case HEADER_NUMBER_CONTAINERS
                Range("D13").value = cleanValue
            Case HEADER_MASTER_UOM
                Range("D15").value = cleanValue
            Case HEADER_QC_STATUS_FLAG
                Range("D16").value = cleanValue
            Case HEADER_LAST_RETEST_DATE
                lastRetestDate = ANSIToDate(cleanValue)
            Case HEADER_FIRST_RETEST_DATE
                firstRetestDate = ANSIToDate(cleanValue)
            Case HEADER_NEW_RETEST
                Range("D19").value = cleanValue
            Case HEADER_QC_TESTED
                Range("D17").value = cleanValue
            Case HEADER_TESTED_DATE
                SetDateStringInCell Range("D18"), ANSIToDate(cleanValue)
        End Select
    Next i
    
    ' Get current balance from RM_Trans.csv and display in QC Balance field
    On Error Resume Next
    Dim currentBalance As Double
    currentBalance = GetQCBalanceFromRMTrans(qcNumber)
    Range("D14").value = currentBalance
    On Error GoTo 0

    ' Handle retest date logic - use last retest if available, otherwise first retest
    Dim retestDateValue As String
    If Len(Trim$(lastRetestDate)) > 0 Then
        retestDateValue = lastRetestDate
    Else
        retestDateValue = firstRetestDate
    End If

    
    ' Set the retest date using the helper function
    If Len(Trim$(retestDateValue)) > 0 Then
        SetDateStringInCell Range("D20"), retestDateValue
    Else
        Range("D20").value = ""
    End If

    GoTo CleanExit

ErrorHandler:
    MsgBox "LoadQCNumberRecord error: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Close #fileNum
    Exit Sub
    
End Sub

Public Sub LoadBatchDetailUsage()
    ' --- Error handling setup ---
    On Error GoTo ErrorHandler

    ' --- Variable declarations ---
    Dim batchNumber As String
    Dim wsBatch As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lastMatchRow As Long
    Dim colBatchNumber As Long
    Dim colPMI As Long
    Dim colDesc As Long
    Dim colExpiry As Long

    ' --- Get the Batch Number from D3 ---
    batchNumber = Trim$(Range("D3").value)
    
    If batchNumber = "" Then
        MsgBox "Please enter a Batch Number in cell D3", vbExclamation, "LoadBatchDetailUsage"
        GoTo CleanExit
    End If
    
    ' --- Object initialization ---
    Set wsBatch = ThisWorkbook.Worksheets("Batch_Details")
    
    ' --- Find the column indices for the required fields ---
    colBatchNumber = FindHeaderCol(wsBatch, 1, FGBNM_HEADER_BATCH_NUMBER)
    colPMI = FindHeaderCol(wsBatch, 1, FGBNM_HEADER_PMI_NUMBER)
    colDesc = FindHeaderCol(wsBatch, 1, FGBNM_HEADER_PRODUCT_DESCRIPTION)
    colExpiry = FindHeaderCol(wsBatch, 1, FGBNM_HEADER_EXPIRY_DATE)
    
    ' --- Validate columns exist ---
    If colBatchNumber = 0 Or colPMI = 0 Or colDesc = 0 Or colExpiry = 0 Then
        MsgBox "Required columns not found in Batch_Details sheet.", vbExclamation, "LoadBatchDetailUsage"
        GoTo CleanExit
    End If
    
    ' --- Search for the LAST occurrence of the batch number ---
    lastRow = wsBatch.Cells(wsBatch.Rows.Count, colBatchNumber).End(xlUp).Row
    lastMatchRow = 0
    
    For r = 2 To lastRow
        If Trim$(wsBatch.Cells(r, colBatchNumber).value) = batchNumber Then
            lastMatchRow = r
        End If
    Next r
    
    ' --- Validate record was found ---
    If lastMatchRow = 0 Then
        MsgBox "Batch Number " & batchNumber & " not found in Batch_Details.", vbExclamation, "LoadBatchDetailUsage"
        Range("D4:D6").ClearContents
        GoTo CleanExit
    End If
    
    ' --- Clear destination cells ---
    Range("D4:D6").ClearContents
    
    ' --- Populate the fields from last matching record ---
    Range("D4").value = wsBatch.Cells(lastMatchRow, colPMI).value
    Range("D5").value = wsBatch.Cells(lastMatchRow, colDesc).value
    Range("D6").NumberFormat = "@"
    Range("D6").value = CStr(wsBatch.Cells(lastMatchRow, colExpiry).value)
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "LoadBatchDetailUsage error: " & Err.Description, vbCritical, "LoadBatchDetailUsage"
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Set wsBatch = Nothing
End Sub

Public Sub PopulateQCNumberList()
    ' --- Error handling setup ---
    On Error GoTo ErrorHandler

    ' --- Variable declarations ---
    Dim rmCode As String
    Dim wsQC As Worksheet
    Dim wsIssuing As Worksheet
    Dim lastRow As Long, r As Long, destRow As Long
    Dim colQCNumber As Long, colRMCode As Long, colRetestDate As Long
    Dim colNilStockFlag As Long, colQCBalance As Long, colMasterUoM As Long, colQCStatus As Long
    Dim colFirstReleaseDate As Long

    ' --- Object initialization ---
    Set wsIssuing = ThisWorkbook.Worksheets("Issuing")
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")

    ' --- Get the RM Code from D10 ---
    rmCode = Trim$(wsIssuing.Range("D10").value)

    ' --- Clear existing list ---
    wsIssuing.Range("H10:M100000").ClearContents

    ' --- Exit if RM Code is blank ---
    If rmCode = "" Then
        GoTo CleanExit
    End If

    ' --- Find required columns in QC_Number sheet ---
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colRMCode = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    colRetestDate = FindHeaderCol(wsQC, 1, HEADER_LAST_RETEST_DATE)
    colNilStockFlag = FindHeaderCol(wsQC, 1, HEADER_NIL_STOCK_FLAG)
    colQCBalance = FindHeaderCol(wsQC, 1, HEADER_QC_BALANCE)
    colMasterUoM = FindHeaderCol(wsQC, 1, HEADER_MASTER_UOM)
    colQCStatus = FindHeaderCol(wsQC, 1, HEADER_QC_STATUS_FLAG)
    colFirstReleaseDate = FindHeaderCol(wsQC, 1, HEADER_FIRST_RELEASE_DATE)

    ' --- Validate columns exist ---
    If colQCNumber = 0 Or colRMCode = 0 Or colRetestDate = 0 Or _
       colNilStockFlag = 0 Or colQCBalance = 0 Or colMasterUoM = 0 Or _
       colQCStatus = 0 Or colFirstReleaseDate = 0 Then
        MsgBox "Required columns not found in QC_Number sheet.", vbExclamation
        GoTo CleanExit
    End If

    ' --- Search for matching RM Code records ---
    lastRow = wsQC.Cells(wsQC.Rows.Count, colRMCode).End(xlUp).Row
    destRow = 10

    For r = 2 To lastRow
        If Trim$(wsQC.Cells(r, colRMCode).value) = rmCode Then
            ' --- Apply filtering criteria ---
            Dim retestDateVal As Variant, nilStockVal As String
            Dim qcBalanceVal As Variant, firstReleaseDateVal As Variant

            retestDateVal = wsQC.Cells(r, colRetestDate).value
            nilStockVal = Trim$(wsQC.Cells(r, colNilStockFlag).value)
            qcBalanceVal = wsQC.Cells(r, colQCBalance).value
            firstReleaseDateVal = wsQC.Cells(r, colFirstReleaseDate).value

            ' --- Filter: Retest Date must be >= Today ---
            If Not IsDate(retestDateVal) Or CDate(retestDateVal) < Date Then
                GoTo NextRecord
            End If

            ' --- Filter: Nil Stock Flag must be "Y" (not "N" or blank) ---
            If nilStockVal = "Y" Then
                GoTo NextRecord
            End If

            ' --- Filter: QC Balance must not be 0 or blank ---
            If qcBalanceVal = 0 Or qcBalanceVal = "" Then
                GoTo NextRecord
            End If

            ' --- Filter: First QC Release Date must not be blank ---
            If firstReleaseDateVal = "" Then
                GoTo NextRecord
            End If

            ' --- Populate the list in columns H through M ---
            wsIssuing.Cells(destRow, "H").value = wsQC.Cells(r, colQCNumber).value
            wsIssuing.Cells(destRow, "I").value = wsQC.Cells(r, colRetestDate).value

            Dim nilStockDisplay As String
            nilStockDisplay = ToYesNo(nilStockVal)
            wsIssuing.Cells(destRow, "J").value = nilStockDisplay

            wsIssuing.Cells(destRow, "K").value = wsQC.Cells(r, colQCBalance).value
            wsIssuing.Cells(destRow, "L").value = wsQC.Cells(r, colMasterUoM).value
            wsIssuing.Cells(destRow, "M").value = wsQC.Cells(r, colQCStatus).value

            destRow = destRow + 1

NextRecord:
        End If
    Next r

    GoTo CleanExit

ErrorHandler:
    MsgBox "Populate QC Number List error: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Set wsQC = Nothing
    Set wsIssuing = Nothing
    Exit Sub

End Sub

Public Sub LoadRMDetailsFromQC()
    ' --- Error handling setup ---
    On Error GoTo ErrorHandler

    ' --- Variable declarations ---
    Dim qcNumber As String
    Dim wsQC As Worksheet
    Dim colQCNumber As Long
    Dim colRMCode As Long
    Dim colStandardName As Long
    Dim colUoM As Long
    Dim lastRow As Long
    Dim r As Long
    Dim foundRecord As Boolean

    ' --- Get the QC Number ---
    qcNumber = GetFieldValByLabelC("QC Number:")

    ' --- Check if D11 is blank ---
    If Trim$(Range("D11").value) = "" Then
        GoTo CleanExit
    End If
    
    ' --- Object initialization ---
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    ' --- Find column indices ---
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colRMCode = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    colStandardName = FindHeaderCol(wsQC, 1, HEADER_STD_NAME)
    colUoM = FindHeaderCol(wsQC, 1, HEADER_MASTER_UOM)
    
    ' --- Validate columns exist ---
    If colQCNumber = 0 Or colRMCode = 0 Or colStandardName = 0 Then
        MsgBox "Required columns not found in QC_Number sheet.", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- Search for the QC Number ---
    lastRow = wsQC.Cells(wsQC.Rows.Count, colQCNumber).End(xlUp).Row
    foundRecord = False
    
    For r = 2 To lastRow
        If Trim$(wsQC.Cells(r, colQCNumber).value) = qcNumber Then
            foundRecord = True
            
            Range("D9:D10").ClearContents
            Range("D9").value = wsQC.Cells(r, colStandardName).value
            Range("D10").value = Format$(wsQC.Cells(r, colRMCode).value, "0000")
            Range("D9:D11").NumberFormat = "@"
            Range("D19").value = wsQC.Cells(r, colUoM).value
            
            Exit For
        End If
    Next r
    
    ' --- Validate record was found ---
    If Not foundRecord Then
        MsgBox "QC Number " & qcNumber & " not found in database.", vbExclamation
        GoTo CleanExit
    End If

    GoTo CleanExit

ErrorHandler:
    MsgBox "LoadRMDetailsFromQC error: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Set wsQC = Nothing
    Exit Sub
    
End Sub

Public Sub ResolveRMCodeFromDescription()
    Dim desc As String
    desc = Trim$(GetFieldValByLabelC("Standard Name:"))
    If desc = "" Then Exit Sub

    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    Dim colStd As Long, colRM As Long
    colStd = FindHeaderCol(wsQC, 1, HEADER_STD_NAME)
    colRM = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    If colStd = 0 Or colRM = 0 Then Exit Sub
    
    Dim lastRow As Long, r As Long
    lastRow = wsQC.Cells(wsQC.Rows.Count, colStd).End(xlUp).Row
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim valStd As String, valRM As String
    For r = 2 To lastRow
        valStd = Trim$(CStr(wsQC.Cells(r, colStd).value))
        If StrComp(valStd, desc, vbTextCompare) = 0 Then
            valRM = Trim$(CStr(wsQC.Cells(r, colRM).value))
            If valRM <> "" Then
                If Not dict.Exists(valRM) Then dict.Add valRM, 1
            End If
        End If
    Next r
    
    If dict.Count = 1 Then
        Range("D10").NumberFormat = "@"
        Range("D10").value = RM_Pad4(dict.Keys()(0))
    Else
        Range("D10").value = "Select from list"
    End If

    GoTo CleanExit

ErrorHandler:
    MsgBox "ResolveRMCodeFromDescription error: " & Err.Description, vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not dict Is Nothing Then Set dict = Nothing
    Set wsQC = Nothing
    Exit Sub
    
End Sub

'-------------------------------------------------------------------------------
' LoadReversalTransactionDetails
'-------------------------------------------------------------------------------
' Purpose: Load specific transaction from RM_Trans.csv by Trans_ID
' Populates form fields D12:D16 with transaction details
'-------------------------------------------------------------------------------
Public Sub LoadReversalTransactionDetails(ByVal transID As String)
    Dim conn As Object
    Dim rs As Object
    Dim fso As Object
    Dim tempFilePath As String
    Dim tempFolder As String
    Dim qcNumber As String
    Dim quantity As String
    Dim masterUoM As String
    Dim userName As String
    Dim transDt As String
    Dim sqlQuery As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    tempFilePath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempFilePath = "" Then
        MsgBox "Unable to create temporary copy of RM_Trans.csv.", vbExclamation, "LoadReversalTransactionDetails"
        GoTo CleanExit
    End If
    
    tempFolder = fso.GetParentFolderName(tempFilePath)
    
    Set conn = OpenADOConnection(tempFolder)
    
    If conn Is Nothing Then
        MsgBox "Unable to establish database connection.", vbCritical, "LoadReversalTransactionDetails"
        GoTo CleanExit
    End If
    
    sqlQuery = "SELECT " & RMT_HEADER_QC_NUMBER & ", " & RMT_HEADER_QUANTITY & ", " & _
               RMT_HEADER_MASTER_UOM & ", " & RMT_HEADER_USERNAME & ", " & RMT_HEADER_TRANS_DT & " " & _
               "FROM [" & fso.GetFileName(tempFilePath) & "] " & _
               "WHERE " & RMT_HEADER_TRANS_ID & " = '" & Replace(transID, "'", "''") & "'"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
        MsgBox "Transaction ID not found: " & transID, vbExclamation, "LoadReversalTransactionDetails"
        GoTo CleanExit
    End If
    
    qcNumber = Trim$(GetFieldValue(rs, RMT_HEADER_QC_NUMBER) & "")
    quantity = Trim$(GetFieldValue(rs, RMT_HEADER_QUANTITY) & "")
    masterUoM = Trim$(GetFieldValue(rs, RMT_HEADER_MASTER_UOM) & "")
    userName = Trim$(GetFieldValue(rs, RMT_HEADER_USERNAME) & "")
    transDt = Trim$(GetFieldValue(rs, RMT_HEADER_TRANS_DT) & "")
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    ' --- Populate form fields ---
    UnprotectActiveSheet
    
    With ActiveSheet
        .Range("D13").value = qcNumber
        .Range("D14").value = quantity
        .Range("D15").value = masterUoM
        .Range("D16").value = userName
        .Range("D17").value = transDt
    End With
    
    ProtectActiveSheet
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "Error in LoadReversalTransactionDetails: " & Err.Description, vbCritical, "LoadReversalTransactionDetails"

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    If tempFilePath <> "" And Not fso Is Nothing Then
        On Error Resume Next
        fso.DeleteFile tempFilePath, True
        On Error GoTo 0
    End If
    Set fso = Nothing
End Sub
'-------------------------------------------------------------------------------
' GetStandardNamesForBatch
'-------------------------------------------------------------------------------
' Purpose: Return array of unique Standard Names issued to a batch
' Used for dropdown population validation
'-------------------------------------------------------------------------------
Public Function GetStandardNamesForBatch(ByVal batchNumber As String) As Variant
    Dim conn As Object
    Dim rs As Object
    Dim fso As Object
    Dim tempFilePath As String
    Dim tempFolder As String
    Dim qcDict As Object
    Dim stdNameDict As Object
    Dim wsQC As Worksheet
    Dim colQCNumber As Long
    Dim colStdName As Long
    Dim lastRow As Long
    Dim r As Long
    Dim qcNumber As String
    Dim stdName As String
    Dim qcKeys As Variant
    Dim j As Long
    Dim sqlQuery As String
    Dim qcVal As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set qcDict = CreateObject("Scripting.Dictionary")
    Set stdNameDict = CreateObject("Scripting.Dictionary")
    
    tempFilePath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempFilePath = "" Then
        GetStandardNamesForBatch = Array()
        GoTo CleanExit
    End If
    
    tempFolder = fso.GetParentFolderName(tempFilePath)
    
    Set conn = OpenADOConnection(tempFolder)
    
    If conn Is Nothing Then
        GetStandardNamesForBatch = Array()
        GoTo CleanExit
    End If
    
    sqlQuery = "SELECT DISTINCT " & RMT_HEADER_QC_NUMBER & " " & _
               "FROM [" & fso.GetFileName(tempFilePath) & "] " & _
               "WHERE " & RMT_HEADER_FG_BATCH & " = '" & Replace(batchNumber, "'", "''") & "' " & _
               "AND " & RMT_HEADER_TRANS_TYPE & " = '" & TRANS_TYPE_CONSUMED & "' " & _
               "AND (" & RMT_HEADER_REVERSED_FLAG & " IS NULL OR " & RMT_HEADER_REVERSED_FLAG & " <> 'Y')"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn
    
    Do While Not rs.EOF
        qcVal = Trim$(GetFieldValue(rs, RMT_HEADER_QC_NUMBER) & "")
        
        If qcVal <> "" And Not qcDict.Exists(qcVal) Then
            qcDict.Add qcVal, True
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    If qcDict.Count = 0 Then
        GetStandardNamesForBatch = Array()
        GoTo CleanExit
    End If
    
    Set wsQC = ThisWorkbook.Worksheets("QC_Numbers")
    
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colStdName = FindHeaderCol(wsQC, 1, HEADER_STANDARD_NAME)
    
    If colQCNumber = 0 Or colStdName = 0 Then
        GetStandardNamesForBatch = Array()
        GoTo CleanExit
    End If
    
    lastRow = wsQC.Cells(wsQC.Rows.Count, colQCNumber).End(xlUp).Row
    qcKeys = qcDict.Keys
    
    For r = 2 To lastRow
        qcNumber = Trim$(CStr(wsQC.Cells(r, colQCNumber).value))
        
        For j = LBound(qcKeys) To UBound(qcKeys)
            If LCase$(qcNumber) = LCase$(qcKeys(j)) Then
                stdName = Trim$(CStr(wsQC.Cells(r, colStdName).value))
                If stdName <> "" And Not stdNameDict.Exists(stdName) Then
                    stdNameDict.Add stdName, True
                End If
            End If
        Next j
    Next r
    
    If stdNameDict.Count = 0 Then
        GetStandardNamesForBatch = Array()
    Else
        GetStandardNamesForBatch = stdNameDict.Keys
    End If
    
    GoTo CleanExit

ErrorHandler:
    GetStandardNamesForBatch = Array()

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    If tempFilePath <> "" And Not fso Is Nothing Then
        fso.DeleteFile tempFilePath, True
    End If
    Set fso = Nothing
    Set qcDict = Nothing
    Set stdNameDict = Nothing
    Set wsQC = Nothing
End Function

'-------------------------------------------------------------------------------
' GetTransactionsForBatchAndRM
'-------------------------------------------------------------------------------
' Purpose: Return 2D array of transactions filtered by batch and RM code
' Excludes reversed transactions. Used for report display.
' Returns: Array with columns (Trans_ID, QC_Number, Quantity, Trans_Dt, Username)
'-------------------------------------------------------------------------------
Public Function GetTransactionsForBatchAndRM(ByVal batchNumber As String, ByVal rmCode As String) As Variant
    Dim conn As Object
    Dim rs As Object
    Dim fso As Object
    Dim tempFilePath As String
    Dim tempFolder As String
    Dim qcToRMDict As Object
    Dim resultList As Object
    Dim wsQC As Worksheet
    Dim colQCNumber As Long
    Dim colRMCode As Long
    Dim lastRow As Long
    Dim r As Long
    Dim qcNum As String
    Dim rmCodeVal As String
    Dim transIDVal As String
    Dim qcVal As String
    Dim quantityVal As String
    Dim transDtVal As String
    Dim usernameVal As String
    Dim resultArray() As Variant
    Dim resultKeys As Variant
    Dim i As Long
    Dim k As Long
    Dim sqlQuery As String
    Dim qcList As String
    Dim qcKeys As Variant
    Dim j As Long
    
    On Error GoTo ErrorHandler
    
    Set qcToRMDict = CreateObject("Scripting.Dictionary")
    Set resultList = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' --- Build QC Number to RM Code mapping from QC_Numbers worksheet ---
    Set wsQC = ThisWorkbook.Worksheets("QC_Number")
    
    colQCNumber = FindHeaderCol(wsQC, 1, HEADER_QC_NUMBER)
    colRMCode = FindHeaderCol(wsQC, 1, HEADER_RM_CODE)
    
    If colQCNumber = 0 Or colRMCode = 0 Then
        GetTransactionsForBatchAndRM = Array()
        GoTo CleanExit
    End If
    
    lastRow = wsQC.Cells(wsQC.Rows.Count, colQCNumber).End(xlUp).Row
    
    For r = 2 To lastRow
        qcNum = Trim$(CStr(wsQC.Cells(r, colQCNumber).value))
        rmCodeVal = Trim$(CStr(wsQC.Cells(r, colRMCode).value))
        
        If qcNum <> "" And LCase$(rmCodeVal) = LCase$(rmCode) Then
            If Not qcToRMDict.Exists(LCase$(qcNum)) Then
                qcToRMDict.Add LCase$(qcNum), True
            End If
        End If
    Next r
    
    If qcToRMDict.Count = 0 Then
        GetTransactionsForBatchAndRM = Array()
        GoTo CleanExit
    End If
    
    tempFilePath = CreateTemporaryFileCopy(FILE_RM_TRANS)
    
    If tempFilePath = "" Then
        GetTransactionsForBatchAndRM = Array()
        GoTo CleanExit
    End If
    
    tempFolder = fso.GetParentFolderName(tempFilePath)
    
    Set conn = OpenADOConnection(tempFolder)
    
    If conn Is Nothing Then
        GetTransactionsForBatchAndRM = Array()
        GoTo CleanExit
    End If
    
    qcKeys = qcToRMDict.Keys
    qcList = ""
    For j = LBound(qcKeys) To UBound(qcKeys)
        If qcList <> "" Then qcList = qcList & ","
        qcList = qcList & "'" & Replace(CStr(qcKeys(j)), "'", "''") & "'"
    Next j
    
    sqlQuery = "SELECT " & RMT_HEADER_TRANS_ID & ", " & RMT_HEADER_QC_NUMBER & ", " & _
               RMT_HEADER_QUANTITY & ", " & RMT_HEADER_TRANS_DT & ", " & RMT_HEADER_USERNAME & " " & _
               "FROM [" & fso.GetFileName(tempFilePath) & "] " & _
               "WHERE " & RMT_HEADER_FG_BATCH & " = '" & Replace(batchNumber, "'", "''") & "' " & _
               "AND " & RMT_HEADER_TRANS_TYPE & " = '" & TRANS_TYPE_CONSUMED & "' " & _
               "AND (" & RMT_HEADER_REVERSED_FLAG & " IS NULL OR " & RMT_HEADER_REVERSED_FLAG & " <> 'Y') " & _
               "AND LCASE(" & RMT_HEADER_QC_NUMBER & ") IN (" & qcList & ")"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn
    
    k = 0
    If Not rs.EOF Then
        Do While Not rs.EOF
            transIDVal = Trim$(GetFieldValue(rs, RMT_HEADER_TRANS_ID) & "")
            qcVal = Trim$(GetFieldValue(rs, RMT_HEADER_QC_NUMBER) & "")
            quantityVal = Trim$(GetFieldValue(rs, RMT_HEADER_QUANTITY) & "")
            transDtVal = Trim$(GetFieldValue(rs, RMT_HEADER_TRANS_DT) & "")
            usernameVal = Trim$(GetFieldValue(rs, RMT_HEADER_USERNAME) & "")
            
            resultList.Add k, transIDVal & "|" & qcVal & "|" & quantityVal & "|" & transDtVal & "|" & usernameVal
            k = k + 1
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    ' --- Build result array ---
    If resultList.Count = 0 Then
        GetTransactionsForBatchAndRM = Array()
        GoTo CleanExit
    End If
    
    ReDim resultArray(1 To resultList.Count, 1 To 5)
    resultKeys = resultList.Keys
    
    For i = 0 To resultList.Count - 1
        Dim parts() As String
        parts = Split(resultList(resultKeys(i)), "|")
        
        resultArray(i + 1, 1) = parts(0)
        resultArray(i + 1, 2) = parts(1)
        resultArray(i + 1, 3) = parts(2)
        resultArray(i + 1, 4) = parts(3)
        resultArray(i + 1, 5) = parts(4)
    Next i
    
    GetTransactionsForBatchAndRM = resultArray
    
    GoTo CleanExit

ErrorHandler:
    GetTransactionsForBatchAndRM = Array()

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    If tempFilePath <> "" And Not fso Is Nothing Then
        On Error Resume Next
        fso.DeleteFile tempFilePath, True
        On Error GoTo 0
    End If
    Set fso = Nothing
    Set qcToRMDict = Nothing
    Set resultList = Nothing
    Set wsQC = Nothing
End Function
