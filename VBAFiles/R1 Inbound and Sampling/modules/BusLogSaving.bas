Attribute VB_Name = "BusLogSaving"
Option Explicit

' Helper function to perform Unit of Measure conversions.
Function ConvertToMasterUoM(ByVal quantity As Double, ByVal enteredUoM As String, ByVal masterUoM As String, ByVal wtPerML As Double) As Double
    Dim convertedQuantity As Double
    convertedQuantity = quantity

    ' Normalize UoM strings to lowercase for comparison
    Dim enteredUoM_lower As String
    Dim masterUoM_lower As String
    enteredUoM_lower = LCase(Trim(enteredUoM))
    masterUoM_lower = LCase(Trim(masterUoM))

    ' --- Conversion Logic ---
    If enteredUoM_lower <> masterUoM_lower Then
        Select Case enteredUoM_lower
            ' Convert Kilograms (kg) to Grams (g) or Milliliters (mL)
            Case "kg", "kgs", "kilogram", "kilograms"
                If masterUoM_lower = "grams" Or masterUoM_lower = "g" Then
                    convertedQuantity = quantity * 1000
                ElseIf masterUoM_lower = "ml" Or masterUoM_lower = "milliliter" Or masterUoM_lower = "milliliters" Then
                    If wtPerML > 0 Then
                        convertedQuantity = (quantity * 1000) / wtPerML
                    End If
                End If
            ' Convert Liters (L) to Grams (g)
            Case "l", "litre", "litres"
                If masterUoM_lower = "grams" Or masterUoM_lower = "g" Then
                    If wtPerML > 0 Then
                        convertedQuantity = quantity * 1000 * wtPerML
                    End If
                ' Convert Liters (L) to Milliliters (mL)
                ElseIf masterUoM_lower = "ml" Or masterUoM_lower = "milliliter" Or masterUoM_lower = "milliliters" Then
                    convertedQuantity = quantity * 1000
                End If
            ' Convert Milliliters (mL) to Grams (g)
            Case "ml", "milliliter", "milliliters"
                If masterUoM_lower = "grams" Or masterUoM_lower = "g" Then
                    If wtPerML > 0 Then
                        convertedQuantity = quantity * wtPerML
                    End If
                End If
            ' Convert Grams (g) to Milliliters (mL)
            Case "g", "gram", "grams"
                 If masterUoM_lower = "ml" Or masterUoM_lower = "milliliter" Or masterUoM_lower = "milliliters" Then
                    If wtPerML > 0 Then
                        convertedQuantity = quantity / wtPerML
                    End If
                 End If
        End Select
    End If

    ConvertToMasterUoM = Round(convertedQuantity, 3)
End Function

' Helper function to validate RM Code requirements
Private Function ValidateRMRequirements(ByVal rmCode As String) As Boolean
    Dim ws As Worksheet
    Dim foundCell As Range
    
    ValidateRMRequirements = False
    
    If rmCode = "" Then Exit Function
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RM_Master")
    If ws Is Nothing Then Exit Function
    On Error GoTo 0
    
    ' Find the RM Code
    Set foundCell = ws.Range("A:A").Find(What:=rmCode, _
                                       LookIn:=xlValues, _
                                       LookAt:=xlWhole, _
                                       SearchOrder:=xlByRows, _
                                       MatchCase:=False)
                                       
    If Not foundCell Is Nothing Then
        ' Check Curr_Ind (column 6) and QC_Number_Required (column 15)
        Dim currInd As String
        Dim qcReq As String
        
        currInd = UCase$(Trim$(ws.Cells(foundCell.row, 6).value))
        qcReq = UCase$(Trim$(ws.Cells(foundCell.row, 15).value))
        
        ValidateRMRequirements = (currInd = "Y" And qcReq = "Y")
    End If
End Function

' Helper function to check if QC Number is unique
Private Function IsQCNumberUnique(ByVal qcNumber As String) As Boolean
    Dim conn As Object
    Dim rs As Object
    Dim folderPath As String
    Dim sql As String
    Dim bomPrefix As String
    
    On Error GoTo CleanupUnique
    
    IsQCNumberUnique = True
    
    If Trim$(qcNumber) = "" Then
        GoTo CleanupUnique
    End If
    
    folderPath = GetWorkbookTablesPath()
    Set conn = OpenADOConnection(folderPath)
    
    If conn Is Nothing Then
        GoTo CleanupUnique
    End If
    
    ' Add BOM character to search value since first column has BOM prefix
    bomPrefix = "ï»¿"
    
    sql = "SELECT COUNT(*) FROM [" & FILE_QC_NUMBER & "] " & _
          "WHERE [" & bomPrefix & HEADER_QC_NUMBER & "] = '" & Replace(qcNumber, "'", "''") & "'"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        IsQCNumberUnique = (rs.Fields(0).value = 0)
    End If

CleanupUnique:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Function

' This is the main macro to be called to save the Goods Receipt data.
Sub SaveGR()

    Dim mainWorkbook As Workbook
    Dim filePath As String
    Dim rmTransFilePath As String
    Dim fso As Object
    Dim ts As Object
    Dim qcBalance As Double
    Dim qcTransID As String
    Dim buomCost As Double
    
    ' Check if files are locked before proceeding
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    rmTransFilePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    If IsFileLocked(filePath) Then
        MsgBox "The file '" & FILE_QC_NUMBER & "' is currently in use by another process. Please close any applications that may be using this file and try again.", vbCritical
        GoTo CleanExit
    End If
    
    If IsFileLocked(rmTransFilePath) Then
        MsgBox "The file '" & FILE_RM_TRANS & "' is currently in use by another process. Please close any applications that may be using this file and try again.", vbCritical
        GoTo CleanExit
    End If
    
    ' Form field variables
    Dim qcNumber As String
    Dim rmCode As String
    Dim StandardName As String
    Dim vendorName As String
    Dim poNumber As String
    Dim mfgName As String
    Dim mfgLot As String
    Dim mfgExpiry As String
    Dim numContainers As Double
    Dim totalQty As Double
    Dim totalUoM As String
    Dim itemCost As Double
    Dim itemCostStr As String
    Dim qcSampleQtyStr As String
    Dim retentionSampleQtyStr As String
    Dim qcSampleQty As Double
    Dim retentionSampleQty As Double
    Dim sampleUoM As String
    Dim quickCode As String
    Dim brandName As String
    Dim partARTG As String
    Dim comments As String
    
    ' RM Master variables
    Dim masterUoM As String
    Dim wtPerML As Double
    Dim packPrice As Double
    Dim rmMasterSheet As Worksheet
    Dim rmCodeRange As Range
    Dim rmTestingFlag As String
    Dim extendedRetestFlag As String

    Set mainWorkbook = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rmMasterSheet = mainWorkbook.Sheets("RM_Master")

    On Error GoTo ErrorHandler
    
    ' Get all form values using GetFieldValByLabelC
    qcNumber = GetFieldValByLabelC("QC Number:")
    rmCode = GetFieldValByLabelC("RM Code:")
    StandardName = GetFieldValByLabelC("Standard Name:")
    vendorName = GetFieldValByLabelC("Vendor Name:")
    poNumber = GetFieldValByLabelC("PO Number:")
    mfgName = GetFieldValByLabelC("Manufacturer Name:")
    mfgLot = GetFieldValByLabelC("Manufacturer Lot Number:")
    mfgExpiry = GetFieldValByLabelC("Manufacturer Expiry:")
    numContainers = Val(GetFieldValByLabelC("No. Containers:"))
    totalQty = Val(GetFieldValByLabelC("Total Quantity:"))
    totalUoM = GetFieldValByLabelC("Unit of Measure:")
    itemCostStr = Trim$(GetFieldValByLabelC("Total Item Cost (excl GST):"))
    
    ' Validate Item Cost is numeric if provided
    If itemCostStr <> "" Then
        If Not IsNumeric(itemCostStr) Then
            MsgBox "Total Item Cost (excl GST) must be a valid number.", vbCritical
            GoTo CleanExit
        End If
        itemCost = CDbl(itemCostStr)
    Else
        itemCost = 0
    End If
    
    qcSampleQtyStr = Trim$(GetFieldValByLabelC("QC Sample QTY:"))
    retentionSampleQtyStr = Trim$(GetFieldValByLabelC("Retention Sample QTY:"))
    sampleUoM = Trim$(GetFieldValByLabelC("Sample UoM:"))
    quickCode = GetFieldValByLabelC("Quick Code:")
    brandName = GetFieldValByLabelC("Brand Name:")
    partARTG = GetFieldValByLabelC("Part Number/ARTG:")
    comments = GetFieldValByLabelC("Comments:")
    
    ' Validate numeric sample fields and convert to Double
    If qcSampleQtyStr <> "" Then
        If Not IsNumeric(qcSampleQtyStr) Then
            MsgBox "QC Sample QTY must be a number.", vbCritical
            GoTo CleanExit
        End If
        qcSampleQty = CDbl(qcSampleQtyStr)
    Else
        qcSampleQty = 0
    End If

    If retentionSampleQtyStr <> "" Then
        If Not IsNumeric(retentionSampleQtyStr) Then
            MsgBox "Retention Sample QTY must be a number.", vbCritical
            GoTo CleanExit
        End If
        retentionSampleQty = CDbl(retentionSampleQtyStr)
    Else
        retentionSampleQty = 0
    End If

    ' If either sample qty was entered, Sample UoM is required
    If (qcSampleQtyStr <> "" Or retentionSampleQtyStr <> "") And sampleUoM = "" Then
        MsgBox "Please specify Sample UoM when entering QC or Retention sample quantities.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate RM Code requirements
    If rmCode <> "" Then
        If Not ValidateRMRequirements(rmCode) Then
            MsgBox "RM Code " & rmCode & " is either not current or does not require QC Numbers.", vbCritical
            GoTo CleanExit
        End If
    End If
    
    ' --- Comma validation for fields being written to file ---
    If Not ValidateNoCommas(vendorName, "Vendor Name") Then GoTo CleanExit
    If Not ValidateNoCommas(poNumber, "PO Number") Then GoTo CleanExit
    If Not ValidateNoCommas(mfgName, "Manufacturer Name") Then GoTo CleanExit
    
    ' Validate required fields
    If qcNumber = "" Or rmCode = "" Or StandardName = "" Or vendorName = "" Or mfgName = "" Or _
       mfgLot = "" Or numContainers = 0 Or totalQty = 0 Or totalUoM = "" Then
        MsgBox "Please fill in all required fields.", vbCritical
        GoTo CleanExit
    End If
    
    ' Validate comment is mandatory when Manufacturer Expiry is blank
    If mfgExpiry = "" And comments = "" Then
        MsgBox "Comments are mandatory when Manufacturer Expiry is not provided.", vbCritical
        GoTo CleanExit
    End If

    ' --- Add validation for numeric fields ---
    If Not IsNumeric(numContainers) Or Val(numContainers) <= 0 Then
        MsgBox "No. Containers must be a positive number.", vbCritical
        GoTo CleanExit
    End If

    If Not IsNumeric(totalQty) Or totalQty <= 0 Then
        MsgBox "Total Quantity must be a positive number.", vbCritical
        GoTo CleanExit
    End If

    If Not fso.FileExists(filePath) Then
        MsgBox "The file '" & FILE_QC_NUMBER & "' was not found in the 'Tables' folder. Macro aborted.", vbCritical
        GoTo CleanExit
    End If

    If RM_COL_MASTERUOM = 0 Then
        InitializeRMMasterColumns
        
        If RM_COL_MASTERUOM = 0 Then
            MsgBox "Failed to initialize RM Master columns. Cannot proceed.", vbCritical
            GoTo CleanExit
        End If
    End If
    
    ' Get UoM, Extended Retest and Wt/mL from RM_Master sheet
    InitializeRMMasterColumns
    Set rmCodeRange = rmMasterSheet.Range("A:A").Find(What:=rmCode, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rmCodeRange Is Nothing Then
        masterUoM = rmMasterSheet.Cells(rmCodeRange.row, RM_COL_MASTERUOM).value
        wtPerML = rmMasterSheet.Cells(rmCodeRange.row, RM_COL_WTPERML).value
        rmTestingFlag = rmMasterSheet.Cells(rmCodeRange.row, RM_COL_TESTINGFLAG).value
        
        ' Get Pack Price from RM Master
        packPrice = Val(rmMasterSheet.Cells(rmCodeRange.row, RM_COL_PACKPRICE).value)
   
        ' Get Extended Retest flag - if blank, default to "N"
        extendedRetestFlag = Trim$(rmMasterSheet.Cells(rmCodeRange.row, RM_COL_EXTENDEDRETEST).value)
        If extendedRetestFlag = "" Then
            extendedRetestFlag = "N"
        End If
        
    Else
        MsgBox "RM_Code not found in RM_Master sheet. Aborting.", vbCritical
        GoTo CleanExit
    End If
    
    ' If sampleUoM is empty, use masterUoM
    If sampleUoM = "" Then
        sampleUoM = masterUoM
    End If
    
    ' Convert quantities to the Master_UoM
    Dim convertedTotalQty As Double
    Dim convertedQCSampleQty As Double
    Dim convertedRetentionSampleQty As Double

    convertedTotalQty = ConvertToMasterUoM(totalQty, totalUoM, masterUoM, wtPerML)
    convertedQCSampleQty = ConvertToMasterUoM(qcSampleQty, sampleUoM, masterUoM, wtPerML)
    convertedRetentionSampleQty = ConvertToMasterUoM(retentionSampleQty, sampleUoM, masterUoM, wtPerML)

    ' Calculate QC_Balance
    qcBalance = convertedTotalQty - (convertedQCSampleQty + convertedRetentionSampleQty)
    If qcBalance <= 0 Then
        MsgBox "Error: QC Balance is zero or negative. Please check your quantities.", vbCritical
        GoTo CleanExit
    End If

    ' Determine QC Status Flag based on Testing Flag and QC Sample Quantity
    Dim qcStatusFlag As String
    If rmTestingFlag = "Y" Then
        If qcSampleQty = 0 Or Trim$(GetFieldValByLabelC("QC Sample QTY:")) = "" Then
            qcStatusFlag = "Receipted"
        Else
            qcStatusFlag = "Sampled"
        End If
    Else
        qcStatusFlag = "Released"
    End If
    
    If Not IsQCNumberUnique(qcNumber) Then
        MsgBox "The QC Number '" & qcNumber & "' already exists. A new QC Number will be created and you will have to resave the record.", vbCritical
        Call GetNextQCNumber
        GoTo CleanExit
    End If

    ' *** GET NEXT QC_TRANS_ID BEFORE WRITING ANY FILES ***
    qcTransID = GetNextQCTransactionID()

    ' Build new line using header constants and GetHeaderIndex
    Dim headers As Variant, values() As String, idx As Integer
    headers = GetQCNumberHeaders
    ReDim values(LBound(headers) To UBound(headers))

    values(GetHeaderIndex(HEADER_QC_NUMBER)) = """" & qcNumber & """"
    values(GetHeaderIndex(HEADER_RM_CODE)) = """" & Format(rmCode, "0000") & """"
    values(GetHeaderIndex(HEADER_STANDARD_NAME)) = """" & StandardName & """"
    values(GetHeaderIndex(HEADER_QUICK_CODE)) = """" & quickCode & """"
    values(GetHeaderIndex(HEADER_BRAND_NAME)) = """" & brandName & """"
    values(GetHeaderIndex(HEADER_PART_ARTG)) = """" & partARTG & """"
    values(GetHeaderIndex(HEADER_VENDOR_NAME)) = """" & vendorName & """"
    values(GetHeaderIndex(HEADER_MANUFACTURER_NAME)) = """" & mfgName & """"
    values(GetHeaderIndex(HEADER_MANUFACTURER_LOT)) = """" & mfgLot & """"
    values(GetHeaderIndex(HEADER_MANUFACTURER_EXPIRY)) = """" & DateToANSI(mfgExpiry) & """"
    values(GetHeaderIndex(HEADER_NUMBER_CONTAINERS)) = """" & numContainers & """"
    values(GetHeaderIndex(HEADER_TOTAL_QTY)) = """" & convertedTotalQty & """"
    values(GetHeaderIndex(HEADER_MASTER_UOM)) = """" & masterUoM & """"
    values(GetHeaderIndex(HEADER_GR_DATE)) = """" & DateToANSI(Format(Date, "dd/mm/yyyy")) & """"
    values(GetHeaderIndex(HEADER_GR_BY)) = """" & Environ("Username") & """"
    values(GetHeaderIndex(HEADER_QC_SAMPLE_QTY)) = """" & Round(convertedQCSampleQty, 3) & """"
    values(GetHeaderIndex(HEADER_RETENTION_SAMPLE_QTY)) = """" & Round(convertedRetentionSampleQty, 3) & """"

    ' Only populate sampled-by/date if a QC or Retention sample qty was entered
    Dim sampledByStr As String, sampleDateStr As String
    If convertedQCSampleQty > 0 Or convertedRetentionSampleQty > 0 Then
        sampledByStr = Environ$("Username")
        sampleDateStr = DateToANSI(Format(Date, "dd/mm/yyyy"))
    Else
        sampledByStr = ""
        sampleDateStr = ""
    End If

    values(GetHeaderIndex(HEADER_SAMPLED_BY)) = """" & sampledByStr & """"
    values(GetHeaderIndex(HEADER_SAMPLE_DATE)) = """" & sampleDateStr & """"
    values(GetHeaderIndex(HEADER_QC_BALANCE)) = """" & Round(qcBalance, 3) & """"
    values(GetHeaderIndex(HEADER_QC_STATUS_FLAG)) = """" & qcStatusFlag & """"
    
    If comments <> "" Then
        values(GetHeaderIndex(HEADER_CHANGE_REASON)) = """Goods Receipt - " & Replace(comments, """", """""") & """"
    Else
        values(GetHeaderIndex(HEADER_CHANGE_REASON)) = """Goods Receipt"""
    End If
    
    values(GetHeaderIndex(HEADER_TRANS_DATE)) = """" & DateToANSI(Format(Date, "dd/mm/yyyy")) & """"
    values(GetHeaderIndex(HEADER_TRANS_USER)) = """" & Environ("Username") & """"
    values(GetHeaderIndex(HEADER_EXTENDED_RETEST)) = """" & extendedRetestFlag & """"
    values(GetHeaderIndex(HEADER_QC_TRANS_ID)) = """" & qcTransID & """"
    values(GetHeaderIndex(HEADER_TESTING_FLAG)) = """" & rmTestingFlag & """"
    values(GetHeaderIndex(HEADER_ITEM_COST)) = """" & itemCost & """"
    values(GetHeaderIndex(HEADER_PO_NUMBER)) = """" & poNumber & """"
    
    ' Calculate BUOM_Cost (per unit cost based on total quantity received)
    If convertedTotalQty > 0 Then
        buomCost = Round(itemCost / convertedTotalQty, 4)
    Else
        buomCost = 0
    End If
    values(GetHeaderIndex(HEADER_BUOM_COST)) = """" & buomCost & """"

    ' Populate release and retest dates if no testing is required
    If rmTestingFlag = "N" Then
        values(GetHeaderIndex(HEADER_FIRST_RELEASE_DATE)) = """" & DateToANSI(Format(Date, "dd/mm/yyyy")) & """"
        values(GetHeaderIndex(HEADER_FIRST_RELEASE_BY)) = """" & Environ("Username") & """"
        values(GetHeaderIndex(HEADER_FIRST_RETEST_DATE)) = """" & DateToANSI(mfgExpiry) & """"
        values(GetHeaderIndex(HEADER_FIRST_RETEST_BY)) = """" & Environ("Username") & """"
        values(GetHeaderIndex(HEADER_LAST_RETEST_DATE)) = """" & DateToANSI(mfgExpiry) & """"
        values(GetHeaderIndex(HEADER_LAST_RETEST_BY)) = """" & Environ("Username") & """"
    End If

    ' *** WRITE TO QC_NUMBER_CHANGES.CSV FIRST ***
    Dim auditPath As String
    auditPath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    Set ts = fso.OpenTextFile(auditPath, 8, True) ' 8 = ForAppending
    ts.WriteLine Join(values, ",")
    ts.Close
    Set ts = Nothing

    ' Write to QC_Numbers.csv
    Set ts = fso.OpenTextFile(filePath, 8, True)
    ts.WriteLine Join(values, ",")
    ts.Close
    Set ts = Nothing
    
    ' Write GRE (Goods Receipt) transaction to RM_Trans with QC_Trans_ID
    Call WriteRMTransaction( _
        transType:=TRANS_TYPE_GOODS_RECEIPT, _
        qcNumber:=qcNumber, _
        qcStat:=qcStatusFlag, _
        quantity:=convertedTotalQty, _
        masterUoM:=masterUoM, _
        qcTransID:=qcTransID _
    )

    ' Write SAM (Sample) transaction if sampling occurred with same QC_Trans_ID
    If convertedQCSampleQty > 0 Or convertedRetentionSampleQty > 0 Then
        Call WriteRMTransaction( _
            transType:=TRANS_TYPE_SAMPLE, _
            qcNumber:=qcNumber, _
            qcStat:=qcStatusFlag, _
            quantity:=-(convertedQCSampleQty + convertedRetentionSampleQty), _
            masterUoM:=masterUoM, _
            qcTransID:=qcTransID _
        )
    End If

    ' --- Refresh data ---
    Call LoadQCNumberData

    MsgBox "Goods Receipt for " & qcNumber & " successfully updated!", vbInformation, "Success"
        
    Call GetNextQCNumber
    
    ClearGR

    GoTo CleanExit

ErrorHandler:
    MsgBox "An unexpected error occurred during goods receipt: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set mainWorkbook = Nothing
    Set rmMasterSheet = Nothing
    Set rmCodeRange = Nothing
End Sub

Sub SaveQCSample()

    Dim mainWorkbook As Workbook
    Dim filePath As String
    Dim rmTransFilePath As String
    Dim fso As Object
    Dim ts As Object
    Dim rng As Range
    Dim cell As Range
    Dim fileContent As String
    Dim lines() As String
    Dim line As String
    Dim lineParts() As String
    Dim found As Boolean
    Dim qcTransID As String

    ' Check if files are locked before proceeding
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    rmTransFilePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    If IsFileLocked(filePath) Then
        MsgBox "The file '" & FILE_QC_NUMBER & "' is currently in use by another process. Please close any applications that may be using this file and try again.", vbCritical
        GoTo CleanExit
    End If
    
    If IsFileLocked(rmTransFilePath) Then
        MsgBox "The file '" & FILE_RM_TRANS & "' is currently in use by another process. Please close any applications that may be using this file and try again.", vbCritical
        GoTo CleanExit
    End If

    Dim qcNumber As String
    Dim rmCode As String
    Dim masterUoM As String
    Dim wtPerML As Double
    Dim qcSampleQty As Double
    Dim retentionSampleQty As Double
    Dim sampleUoM As String
    Dim convertedQCSampleQty As Double
    Dim convertedRetentionSampleQty As Double
    Dim newTotalSampleQty As Double
    Dim newQCBalance As Double
    Dim rmMasterSheet As Worksheet
    Dim rmCodeRange As Range
    Dim i As Long
    Dim totalDraw As Double

    Set mainWorkbook = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rmMasterSheet = mainWorkbook.Sheets("RM_Master")

    On Error GoTo ErrorHandler

    ' --- Get values from input sheet using the helper function ---
    qcNumber = GetFieldValByLabelC("QC Number:")
    rmCode = GetFieldValByLabelC("RM Code:")
    qcSampleQty = Val(GetFieldValByLabelC("QC Sample QTY:"))
    retentionSampleQty = Val(GetFieldValByLabelC("Retention Sample QTY:"))
    sampleUoM = Trim$(GetFieldValByLabelC("Sample UoM:"))

    ' --- Validate essential fields ---
    If qcNumber = "" Or rmCode = "" Or qcSampleQty = 0 Or sampleUoM = "" Then
        MsgBox "Please fill in all required fields (QC Number, RM Code, QC Sample QTY, and Sample UoM).", vbCritical
        GoTo CleanExit
    End If

    If Not fso.FileExists(filePath) Then
        MsgBox "The file '" & FILE_QC_NUMBER & "' was not found in the 'Tables' folder. Macro aborted.", vbCritical
        GoTo CleanExit
    End If

    ' Validate RM Code requirements
    If Not ValidateRMRequirements(rmCode) Then
        MsgBox "RM Code " & rmCode & " is either not current or does not require QC Numbers.", vbCritical
        GoTo CleanExit
    End If

    ' --- Get UoM and Wt/mL from RM_Master sheet ---
    InitializeRMMasterColumns
    Set rmCodeRange = rmMasterSheet.Range("A:A").Find(What:=rmCode, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rmCodeRange Is Nothing Then
        masterUoM = rmMasterSheet.Cells(rmCodeRange.row, RM_COL_MASTERUOM).value
        wtPerML = rmMasterSheet.Cells(rmCodeRange.row, RM_COL_WTPERML).value
    Else
        MsgBox "RM_Code not found in RM_Master sheet. Aborting.", vbCritical
        GoTo CleanExit
    End If
    
    ' --- Convert quantities to the Master_UoM ---
    convertedQCSampleQty = ConvertToMasterUoM(qcSampleQty, sampleUoM, masterUoM, wtPerML)
    convertedRetentionSampleQty = ConvertToMasterUoM(retentionSampleQty, sampleUoM, masterUoM, wtPerML)
    totalDraw = convertedQCSampleQty + convertedRetentionSampleQty

    ' --- Check current balance from RM_Trans ---
    Dim currentBalance As Double
    currentBalance = GetQCBalanceFromRMTrans(qcNumber)
    
    If totalDraw > currentBalance Then
        MsgBox "Sample quantity (" & Format(totalDraw, "0.00") & " " & masterUoM & _
               ") exceeds current balance (" & Format(currentBalance, "0.00") & " " & masterUoM & ")." & vbCrLf & _
               "Operation cancelled.", vbCritical
        GoTo CleanExit
    End If
    
    ' --- Handle exact balance exhaustion ---
    Dim nilStockUpdate As Boolean
    nilStockUpdate = False
    
    If totalDraw = currentBalance Then
        Dim response As VbMsgBoxResult
        response = MsgBox("This sample will exactly exhaust the remaining balance (" & Format(currentBalance, "0.00") & " " & masterUoM & ")." & vbCrLf & _
                         "The item will be marked as Nil Stock." & vbCrLf & vbCrLf & _
                         "Do you want to proceed?", vbYesNo + vbExclamation, "Nil Stock Warning")
        
        If response = vbNo Then
            GoTo CleanExit
        End If
        
        nilStockUpdate = True
    End If

    ' *** GET NEXT QC_TRANS_ID BEFORE WRITING ANY FILES ***
    qcTransID = GetNextQCTransactionID()

    ' --- Read CSV and find the matching QC Number entry ---
    Set ts = fso.OpenTextFile(filePath, 1, False)
    fileContent = ts.ReadAll
    ts.Close

    lines = Split(fileContent, vbCrLf)
    found = False

    Dim idxQC As Long, idxTotSample As Long, idxQCBal As Long
    Dim idxQCSample As Long, idxRetSample As Long, idxSampleBy As Long, idxSampleDate As Long, idxStatus As Long
    Dim idxNilStockFlag As Long, idxNilStockFlagDate As Long, idxNilStockFlagBy As Long
    Dim idxQCTransID As Long, idxChangeReason As Long
    
    idxQC = GetHeaderIndex0(HEADER_QC_NUMBER)
    idxTotSample = GetHeaderIndex0(HEADER_TOTAL_SAMPLE_QTY)
    idxQCBal = GetHeaderIndex0(HEADER_QC_BALANCE)
    idxQCSample = GetHeaderIndex0(HEADER_QC_SAMPLE_QTY)
    idxRetSample = GetHeaderIndex0(HEADER_RETENTION_SAMPLE_QTY)
    idxSampleBy = GetHeaderIndex0(HEADER_SAMPLED_BY)
    idxSampleDate = GetHeaderIndex0(HEADER_SAMPLE_DATE)
    idxStatus = GetHeaderIndex0(HEADER_QC_STATUS_FLAG)
    idxNilStockFlag = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG)
    idxNilStockFlagDate = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_DATE)
    idxNilStockFlagBy = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_BY)
    idxQCTransID = GetHeaderIndex0(HEADER_QC_TRANS_ID)
    idxChangeReason = GetHeaderIndex0(HEADER_CHANGE_REASON)

    For i = 1 To UBound(lines)
        line = lines(i)
        If Trim(line) <> "" Then
            lineParts = Split(line, ",")

            If UBound(lineParts) >= idxQC And _
               LCase$(Trim$(Replace(lineParts(idxQC), """", ""))) = LCase$(Trim$(qcNumber)) Then

                ' Get existing quantities from CSV
                Dim existingTotalSampleQty As Double
                Dim existingQCSampleQty As Double
                Dim existingRetentionSampleQty As Double

                existingTotalSampleQty = CDbl(Val(Replace(lineParts(idxTotSample), """", "")))
                existingQCSampleQty = CDbl(Val(Replace(lineParts(idxQCSample), """", "")))
                existingRetentionSampleQty = CDbl(Val(Replace(lineParts(idxRetSample), """", "")))
                
                ' Calculate new quantities
                newTotalSampleQty = existingTotalSampleQty + totalDraw
                newQCBalance = currentBalance - totalDraw

                ' Update all sample quantities
                lineParts(idxTotSample) = """" & Round(newTotalSampleQty, 3) & """"
                lineParts(idxQCSample) = """" & Round(existingQCSampleQty + convertedQCSampleQty, 3) & """"
                lineParts(idxRetSample) = """" & Round(existingRetentionSampleQty + convertedRetentionSampleQty, 3) & """"
                lineParts(idxQCBal) = """" & Round(newQCBalance, 3) & """"
                
                ' Update status and audit fields
                lineParts(idxSampleBy) = """" & Environ$("Username") & """"
                lineParts(idxSampleDate) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"
                lineParts(idxStatus) = """Sampled"""
                lineParts(GetHeaderIndex0(HEADER_TRANS_DATE)) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"
                lineParts(GetHeaderIndex0(HEADER_TRANS_USER)) = """" & Environ$("Username") & """"
                lineParts(idxQCTransID) = """" & qcTransID & """"
                lineParts(idxChangeReason) = """Sampled"""
                
                ' Update Nil Stock fields if balance is exhausted
                If nilStockUpdate Then
                    lineParts(idxNilStockFlag) = """Y"""
                    lineParts(idxNilStockFlagDate) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"
                    lineParts(idxNilStockFlagBy) = """" & Environ$("Username") & """"
                End If

                lines(i) = Join(lineParts, ",")
                found = True
                Exit For
            End If
        End If
    Next i

    If Not found Then
        MsgBox "QC Number '" & qcNumber & "' not found in QC_Number.csv. No entry was updated.", vbExclamation
        GoTo CleanExit
    End If
    
    ' *** WRITE TO QC_NUMBER_CHANGES.CSV FIRST (AUDIT TRAIL) ***
    Dim auditPath As String
    auditPath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    Set ts = fso.OpenTextFile(auditPath, 8, True) ' 8 = ForAppending
    ts.WriteLine lines(i)
    ts.Close
    Set ts = Nothing
    
    ' --- Overwrite the QC_Number.csv file with the updated content ---
    Set ts = fso.OpenTextFile(filePath, 2, False) ' 2 = ForWriting
    ts.Write Join(lines, vbCrLf)
    ts.Close
    Set ts = Nothing
    
    ' Write SAM (Sample) transaction to RM_Trans with QC_Trans_ID
    Call WriteRMTransaction( _
        transType:=TRANS_TYPE_SAMPLE, _
        qcNumber:=qcNumber, _
        qcStat:="Sampled", _
        quantity:=-(convertedQCSampleQty + convertedRetentionSampleQty), _
        masterUoM:=masterUoM, _
        qcTransID:=qcTransID _
    )
    
    ' --- Refresh data ---
    Call LoadQCNumberData
    
    ' --- Clear the input cells ---
    ClearSample
    MsgBox "QC Sample for " & qcNumber & " successfully updated!", vbInformation, "Success"

    GoTo CleanExit

' --- Error Handling ---
ErrorHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set mainWorkbook = Nothing
    Set fso = Nothing
    Set ts = Nothing
    Set rng = Nothing
    Set cell = Nothing
    Set rmMasterSheet = Nothing
    Set rmCodeRange = Nothing
End Sub

Private Function ValidateNoCommas(ByVal value As String, ByVal fieldName As String) As Boolean
    If InStr(1, value, ",") > 0 Then
        MsgBox "Error: " & fieldName & " cannot contain commas. Please remove any commas and try again.", vbCritical
        ValidateNoCommas = False
    Else
        ValidateNoCommas = True
    End If
End Function

Public Sub SaveVendorList()

    Dim wsMaintain As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim lastRow As Long
    Dim r As Long, outRow As Long
    Dim vendorName As String
    Dim dataRange As Range
    Dim headerValue As String
    
    On Error GoTo ErrorHandler
    
    ' Get the Maintain Vendor worksheet
    Set wsMaintain = ThisWorkbook.Worksheets("Maintain Vendor")
    
    ' Get the destination file path
    filePath = BuildFilePath(ThisWorkbook.Path, FOLDER_TABLES, FILE_VENDOR_LIST)
    
    ' Find last row with data (starting from B5)
    lastRow = wsMaintain.Cells(10000, 2).End(xlUp).row
    
    ' Validate no commas in data
    For r = 5 To lastRow
        vendorName = Trim(wsMaintain.Cells(r, 2).value)
        If vendorName <> "" Then
            If Not ValidateNoCommas(vendorName, "Vendor Name") Then
                GoTo CleanExit
            End If
        End If
    Next r
    
    ' Sort the data (excluding header)
    If lastRow > 4 Then
        Set dataRange = wsMaintain.Range("B5:B" & lastRow)
        With wsMaintain.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dataRange, Order:=xlAscending
            .SetRange dataRange
            .Header = xlNo
            .Apply
        End With
    End If
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open file for writing (overwrite existing)
    Set ts = fso.OpenTextFile(filePath, 2, True)
    
    ' Write fixed header
    ts.WriteLine """Vendor_Name"""
    
    ' Write data rows starting at row 2 in destination
    For r = 5 To lastRow  ' Start from row 5 (B5)
        vendorName = Trim(wsMaintain.Cells(r, 2).value)
        If vendorName <> "" Then
            ts.WriteLine """" & UCase$(vendorName) & """"
        End If
    Next r
    
    ' Close file
    ts.Close
    
    ' Reload the vendor list to update all references
    Call LoadVendorList
    
    MsgBox "Vendor List has been successfully updated!", vbInformation
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error saving Vendor List: " & Err.Description, vbCritical

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set wsMaintain = Nothing
    Set dataRange = Nothing
End Sub

Public Sub SaveManufacturerList()

    Dim wsMaintain As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim lastRow As Long
    Dim r As Long, outRow As Long
    Dim manufacturerName As String
    Dim dataRange As Range
    Dim headerValue As String
    
    On Error GoTo ErrorHandler
    
    ' Get the Maintain Manufacturer worksheet
    Set wsMaintain = ThisWorkbook.Worksheets("Maintain Manufacturer")
    
    ' Get the destination file path
    filePath = BuildFilePath(ThisWorkbook.Path, FOLDER_TABLES, FILE_MANUFACTURER_LIST)
    
    ' Find last row with data (starting from B5)
    lastRow = wsMaintain.Cells(10000, 2).End(xlUp).row
    
    ' Validate no commas in data
    For r = 5 To lastRow
        manufacturerName = Trim(wsMaintain.Cells(r, 2).value)
        If manufacturerName <> "" Then
            If Not ValidateNoCommas(manufacturerName, "Manufacturer Name") Then
                GoTo CleanExit
            End If
        End If
    Next r
    
    ' Sort the data (excluding header)
    If lastRow > 4 Then
        Set dataRange = wsMaintain.Range("B5:B" & lastRow)
        With wsMaintain.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dataRange, Order:=xlAscending
            .SetRange dataRange
            .Header = xlNo
            .Apply
        End With
    End If
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open file for writing (overwrite existing)
    Set ts = fso.OpenTextFile(filePath, 2, True)
    
    ' Write fixed header
    ts.WriteLine """Manufacturer_Name"""
    
    ' Write data rows starting at row 2 in destination
    For r = 5 To lastRow  ' Start from row 5 (B5)
        manufacturerName = Trim(wsMaintain.Cells(r, 2).value)
        If manufacturerName <> "" Then
            ts.WriteLine """" & UCase$(manufacturerName) & """"
        End If
    Next r
    
    ' Close file
    ts.Close
    
    ' Reload the manufacturer list to update all references
    Call LoadManufacturerList
    
    MsgBox "Manufacturer List has been successfully updated!", vbInformation
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error saving Manufacturer List: " & Err.Description, vbCritical

CleanExit:
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set wsMaintain = Nothing
    Set dataRange = Nothing
End Sub


