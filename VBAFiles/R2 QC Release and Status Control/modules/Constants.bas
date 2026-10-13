Attribute VB_Name = "Constants"
Option Explicit

' --- Master Data Root Path (use this as the base for Master Data files) ---
' Replaced single constant with per-environment roots and a simple switch.
Public Const DEV_MASTER_DATA_ROOT As String = "G:\Pharmacy\2.MFG\DEV\Master Data\"
Public Const PROD_MASTER_DATA_ROOT As String = "G:\Pharmacy\2.MFG\DATA\Master Data\"
Public Const HOME_MASTER_DATA_ROOT As String = "C:\Pharmacy\2.MFG\DATA\Master Data\"

' Master Switches
Public Const MASTER_ENV_SWITCH As String = "PROD"
Public Const PROD_DEV As String = "PROD" ' DEV or PROD controls backups, form protection, cell dragging and save as suppression

' --- File and Folder Names ---
Public Const FOLDER_TABLES As String = "Tables"
Public Const FOLDER_EXPORTS As String = "Exports"

' Files used by workbook-local Tables
Public Const FILE_QC_NUMBER As String = "QC_Number.csv"
Public Const FILE_QC_NUMBER_CHANGES As String = "QC_Number_Changes.csv"
Public Const FILE_RM_TRANS As String = "RM_Trans.csv"

' --- RM_Trans.csv Headers ---
Public Const RMT_HEADER_TRANS_ID As String = "RM_Trans_ID"
Public Const RMT_HEADER_TRANS_TYPE As String = "Trans_Type"
Public Const RMT_HEADER_QC_NUMBER As String = "QC_Number"
Public Const RMT_HEADER_QC_STAT As String = "QC_Stat"
Public Const RMT_HEADER_FG_BATCH As String = "FG_Batch"
Public Const RMT_HEADER_FG_BATCH_STATUS As String = "FG_Batch_Status"
Public Const RMT_HEADER_QUANTITY As String = "Quantity"
Public Const RMT_HEADER_MASTER_UOM As String = "Master_UoM"
Public Const RMT_HEADER_USERNAME As String = "Username"
Public Const RMT_HEADER_TRANS_DT As String = "Trans_Dt"
Public Const RMT_HEADER_QC_TRANS_ID As String = "QC_Trans_ID"
Public Const RMT_HEADER_PI_TRANS_ID As String = "PI_Trans_ID"
Public Const RMT_HEADER_ADJ_NOTE As String = "Adj_Note"
Public Const RMT_HEADER_REVERSED_FLAG As String = "Reversed_Flag"
Public Const RMT_HEADER_MEASURED_BY As String = "Measured_By"

' Transaction Types for RM_Trans.csv
Public Const TRANS_TYPE_GOODS_RECEIPT As String = "GRE"
Public Const TRANS_TYPE_SAMPLE As String = "SAM"
Public Const TRANS_TYPE_CONSUMED As String = "CON"
Public Const TRANS_TYPE_WRITE_ON As String = "WON"
Public Const TRANS_TYPE_WRITE_OFF As String = "WOF"
Public Const TRANS_TYPE_GOODS_RECEIPT_REVERSAL As String = "GRR"
Public Const TRANS_TYPE_SAMPLING_REVERSAL As String = "SAR"
Public Const TRANS_TYPE_REJECTED As String = "REJ"

' QC Status Constants
Public Const QC_STATUS_REJECTED As String = "Rejected"
Public Const QC_STATUS_CANCELLED As String = "Cancelled"
Public Const QC_STATUS_RELEASED As String = "Released"
Public Const QC_STATUS_QUARANTINE As String = "Quarantine"
Public Const QC_STATUS_SAMPLED As String = "Sampled"
Public Const QC_STATUS_RECEIPTED As String = "Receipted"

' Trans_ID Formatting Constants
Public Const TRANS_ID_FORMAT_LENGTH As Long = 6
Public Const TRANS_ID_START_VALUE As Long = 1
Public Const TRANS_ID_PADDING_CHAR As String = "0"
Public Const TRANS_ID_FORMAT_STRING As String = "000000"

' File Extension Constants
Public Const EXCEL_FILE_EXT As String = ".xlsx"
Public Const CSV_FILE_EXT As String = ".csv"

' File Prefix Constants
Public Const PREFIX_HEADERS As String = "QC_Numbers_"
Public Const PREFIX_CHANGES As String = "QC_Number_Changes_"

' --- Error Messages ---
Public Const ERROR_HEADER_MISSING As String = "One or more required headers were not found in the source CSV."
Public Const ERROR_CSV_IMPORT As String = "Could not import CSV file: "
Public Const ERROR_INVALID_DATA As String = "Invalid data detected in: "

' --- QC_Number.csv Headers ---
Public Const HEADER_QC_NUMBER As String = "QC_Number"
Public Const HEADER_RM_CODE As String = "RM_Code"
Public Const HEADER_STANDARD_NAME As String = "Standard_Name"
Public Const HEADER_QUICK_CODE As String = "Quick_Code"
Public Const HEADER_BRAND_NAME As String = "Brand_Name"
Public Const HEADER_PART_ARTG As String = "Part_ARTG"
Public Const HEADER_VENDOR_NAME As String = "Vendor_Name"
Public Const HEADER_MANUFACTURER_NAME As String = "Mfg_Name"
Public Const HEADER_MANUFACTURER_LOT As String = "Mfg_Lot"
Public Const HEADER_MANUFACTURER_EXPIRY As String = "Mfg_Expiry"
Public Const HEADER_NUMBER_CONTAINERS As String = "Number_Containers"
Public Const HEADER_TOTAL_QTY As String = "Total_Qty"
Public Const HEADER_TOTAL_SAMPLE_QTY As String = "Total_Sample_Qty"
Public Const HEADER_MASTER_UOM As String = "Master_UoM"
Public Const HEADER_GR_DATE As String = "GR_Date"
Public Const HEADER_GR_BY As String = "GR_By"
Public Const HEADER_QC_SAMPLE_QTY As String = "QC_Sample_Qty"
Public Const HEADER_RETENTION_SAMPLE_QTY As String = "Retention_Sample_Qty"
Public Const HEADER_SAMPLED_BY As String = "Sampled_By"
Public Const HEADER_SAMPLE_DATE As String = "Sample_Date"
Public Const HEADER_QC_BALANCE As String = "QC_Balance"
Public Const HEADER_TESTING_FLAG As String = "Testing_Flag"
Public Const HEADER_FIRST_RELEASE_DATE As String = "First_Release_Date"
Public Const HEADER_FIRST_RELEASE_BY As String = "First_Release_by"
Public Const HEADER_LAST_RELEASE_DATE As String = "Last_Release_Date"
Public Const HEADER_LAST_RELEASE_BY As String = "Last_Release_By"
Public Const HEADER_FIRST_RETEST_DATE As String = "First_Retest_Date"
Public Const HEADER_FIRST_RETEST_BY As String = "First_Retest_By"
Public Const HEADER_LAST_RETEST_DATE As String = "Last_Retest_Date"
Public Const HEADER_LAST_RETEST_BY As String = "Last_Retest_By"
Public Const HEADER_NIL_STOCK_FLAG As String = "Nil_Stock_Flag"
Public Const HEADER_NIL_STOCK_FLAG_DATE As String = "Nil_Stock_Flag_Date"
Public Const HEADER_NIL_STOCK_FLAG_BY As String = "Nil_Stock_Flag_By"
Public Const HEADER_QC_STATUS_FLAG As String = "QC_Status_Flag"
Public Const HEADER_CHANGE_REASON As String = "Change_Reason"
Public Const HEADER_TRANS_DATE As String = "Trans_Date"
Public Const HEADER_TRANS_USER As String = "Trans_User"
Public Const HEADER_NEW_RETEST As String = "New_Retest"
Public Const HEADER_QC_TESTED As String = "QC_Tested"
Public Const HEADER_TESTED_DATE As String = "Tested_Date"
Public Const HEADER_EXTENDED_RETEST As String = "Extended_Retest"
Public Const HEADER_QC_TRANS_ID As String = "QC_Trans_ID"
' --- Helper to resolve master root based on switch / auto-detect ---
Private Function ResolveMasterDataRoot() As String
    Dim candidate As String
    Select Case UCase$(Trim$(MASTER_ENV_SWITCH))
        Case "DEV"
            ResolveMasterDataRoot = DEV_MASTER_DATA_ROOT
            Exit Function
        Case "PROD"
            ResolveMasterDataRoot = PROD_MASTER_DATA_ROOT
            Exit Function
        Case "HOME"
            ResolveMasterDataRoot = HOME_MASTER_DATA_ROOT
            Exit Function
        Case Else ' AUTO or unknown -> try to detect DEV, then PROD, then HOME
            candidate = DEV_MASTER_DATA_ROOT
            If FolderExists(candidate) Then
                ResolveMasterDataRoot = candidate
                Exit Function
            End If
            candidate = PROD_MASTER_DATA_ROOT
            If FolderExists(candidate) Then
                ResolveMasterDataRoot = candidate
                Exit Function
            End If
            ResolveMasterDataRoot = HOME_MASTER_DATA_ROOT
            Exit Function
    End Select
End Function

' --- Helper to build full path under MASTER_DATA_ROOT ---
Public Function GetMasterDataPath(Optional ByVal relativePath As String = "") As String
    Dim base As String
    base = ResolveMasterDataRoot()
    If Len(relativePath) = 0 Then
        GetMasterDataPath = base
        Exit Function
    End If
    If Right(base, 1) <> "\" Then base = base & "\"
    If Left(relativePath, 1) = "\" Then relativePath = Mid$(relativePath, 2)
    GetMasterDataPath = base & relativePath
End Function

' --- Helper: workbook-local Tables folder path (based on workbook location) ---
Public Function GetWorkbookTablesPath() As String
    Dim base As String
    base = ThisWorkbook.Path
    If Len(base) = 0 Then
        ' Workbook unsaved — return current folder marker
        GetWorkbookTablesPath = FOLDER_TABLES & "\"
        Exit Function
    End If
    If Right(base, 1) <> "\" Then base = base & "\"
    GetWorkbookTablesPath = base & FOLDER_TABLES & "\"
End Function

' --- Helper: full path to a file in the workbook's Tables folder ---
Public Function GetWorkbookTableFilePath(ByVal fileName As String) As String
    Dim tblPath As String
    tblPath = GetWorkbookTablesPath()
    If Left(fileName, 1) = "\" Then fileName = Mid$(fileName, 2)
    GetWorkbookTableFilePath = tblPath & fileName
End Function

' --- Utility: Array of All QC_Number Headers ---
Public Function GetQCNumberHeaders() As Variant
    Dim headers(1 To 42) As String

    headers(1) = HEADER_QC_NUMBER
    headers(2) = HEADER_RM_CODE
    headers(3) = HEADER_STANDARD_NAME
    headers(4) = HEADER_QUICK_CODE
    headers(5) = HEADER_BRAND_NAME
    headers(6) = HEADER_PART_ARTG
    headers(7) = HEADER_VENDOR_NAME
    headers(8) = HEADER_MANUFACTURER_NAME
    headers(9) = HEADER_MANUFACTURER_LOT
    headers(10) = HEADER_MANUFACTURER_EXPIRY
    headers(11) = HEADER_NUMBER_CONTAINERS
    headers(12) = HEADER_TOTAL_QTY
    headers(13) = HEADER_TOTAL_SAMPLE_QTY
    headers(14) = HEADER_MASTER_UOM
    headers(15) = HEADER_GR_DATE
    headers(16) = HEADER_GR_BY
    headers(17) = HEADER_QC_SAMPLE_QTY
    headers(18) = HEADER_RETENTION_SAMPLE_QTY
    headers(19) = HEADER_SAMPLED_BY
    headers(20) = HEADER_SAMPLE_DATE
    headers(21) = HEADER_QC_BALANCE
    headers(22) = HEADER_TESTING_FLAG
    headers(23) = HEADER_FIRST_RELEASE_DATE
    headers(24) = HEADER_FIRST_RELEASE_BY
    headers(25) = HEADER_LAST_RELEASE_DATE
    headers(26) = HEADER_LAST_RELEASE_BY
    headers(27) = HEADER_FIRST_RETEST_DATE
    headers(28) = HEADER_FIRST_RETEST_BY
    headers(29) = HEADER_LAST_RETEST_DATE
    headers(30) = HEADER_LAST_RETEST_BY
    headers(31) = HEADER_NIL_STOCK_FLAG
    headers(32) = HEADER_NIL_STOCK_FLAG_DATE
    headers(33) = HEADER_NIL_STOCK_FLAG_BY
    headers(34) = HEADER_QC_STATUS_FLAG
    headers(35) = HEADER_CHANGE_REASON
    headers(36) = HEADER_TRANS_DATE
    headers(37) = HEADER_TRANS_USER
    headers(38) = HEADER_NEW_RETEST
    headers(39) = HEADER_QC_TESTED
    headers(40) = HEADER_TESTED_DATE
    headers(41) = HEADER_EXTENDED_RETEST
    headers(42) = HEADER_QC_TRANS_ID

    GetQCNumberHeaders = headers
End Function

' --- Utility: Array of All RM_Trans Headers ---
Public Function GetRMTransHeaders() As Variant
     Dim headers(1 To 15) As String
    
    headers(1) = RMT_HEADER_TRANS_ID
    headers(2) = RMT_HEADER_TRANS_TYPE
    headers(3) = RMT_HEADER_QC_NUMBER
    headers(4) = RMT_HEADER_QC_STAT
    headers(5) = RMT_HEADER_FG_BATCH
    headers(6) = RMT_HEADER_FG_BATCH_STATUS
    headers(7) = RMT_HEADER_QUANTITY
    headers(8) = RMT_HEADER_MASTER_UOM
    headers(9) = RMT_HEADER_USERNAME
    headers(10) = RMT_HEADER_TRANS_DT
    headers(11) = RMT_HEADER_QC_TRANS_ID
    headers(12) = RMT_HEADER_PI_TRANS_ID
    headers(13) = RMT_HEADER_ADJ_NOTE
    headers(14) = RMT_HEADER_REVERSED_FLAG
    headers(15) = RMT_HEADER_MEASURED_BY
    
    GetRMTransHeaders = headers
End Function