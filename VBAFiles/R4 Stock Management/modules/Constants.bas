Attribute VB_Name = "Constants"
Option Explicit

' --- Master Data Root Path (use this as the base for Master Data files) ---
' Replaced single constant with per-environment roots and a simple switch.
Public Const DEV_MASTER_DATA_ROOT As String = "G:\Pharmacy\2.MFG\DEV\Master Data\"
Public Const PROD_MASTER_DATA_ROOT As String = "G:\Pharmacy\2.MFG\DATA\Master Data\"
Public Const HOME_MASTER_DATA_ROOT As String = "C:\Pharmacy\2.MFG\DATA\Master Data\"

' Master Switches
Public Const MASTER_ENV_SWITCH As String = "PROD"
Public Const PROD_DEV As String = "PROD"

' --- Folder Names ---
Public Const FOLDER_TABLES As String = "Tables"
Public Const FOLDER_MASTER_DATA As String = "Master Data"
Public Const FOLDER_MATERIAL_MANAGEMENT As String = "Material Management"

' --- Master Data Base Path (relative to workbook parent) ---
Public Const PATH_MASTER_DATA_TABLES As String = "Master Data\Tables"
Public Const PATH_MATERIAL_MANAGEMENT_TABLES As String = "Material Management\Tables"

' Data tables paths
Public Const MASTER_DATA_PATH As String = "\Pharmacy\2.MFG\DATA\Master Data\"
Public Const MATERIAL_MANAGEMENT_PATH As String = "\Pharmacy\2.MFG\DATA\Material Management\"

' Files used by workbook-local Tables
' Files are UTF8 CSVs with BOM encoding
Public Const FILE_QC_NUMBER As String = "QC_Number.csv"
Public Const FILE_RM_TRANS As String = "RM_Trans.csv"
Public Const FILE_BATCH_DETAILS As String = "Label_Prints.csv" ' contains current batch details
Public Const FILE_QC_NUMBER_CHANGES As String = "QC_Number_Changes.csv"
Public Const FILE_RM_MASTER As String = "Raw_Material_Master.csv"
Public Const FILE_PHYSICAL_INVENTORY As String = "Physical_Inventory.csv"
Public Const FILE_LAST_IDS As String = "Last_IDs.csv"

' File Extension Constants
Public Const EXCEL_FILE_EXT As String = ".xlsx"
Public Const CSV_FILE_EXT As String = ".csv"

' File Prefix Constants
Public Const PREFIX_QC_NUMBERS As String = "QC_Numbers_"
Public Const PREFIX_QC_CHANGES As String = "QC_Number_Changes_"

' --- Error Messages ---
Public Const ERROR_HEADER_MISSING As String = "One or more required headers were not found in the source CSV."
Public Const ERROR_CSV_IMPORT As String = "Could not import CSV file: "
Public Const ERROR_INVALID_DATA As String = "Invalid data detected in: "

' Index file ID types
Public Const IDX_TYPE_QC_NUMBER As String = "QC_Number"
Public Const IDX_TYPE_QC_TRANS_ID As String = "QC_Trans_ID"
Public Const IDX_TYPE_RM_TRANS_ID As String = "RM_Trans_ID"
Public Const IDX_TYPE_PI_DOC_NUMBER As String = "PI_Doc_Number"

' Index file headers
Public Const IDX_HEADER_ID_TYPE As String = "ID_Type"
Public Const IDX_HEADER_LAST_VALUE As String = "Last_Value"
Public Const IDX_HEADER_LAST_UPDATED As String = "Last_Updated"
Public Const IDX_HEADER_UPDATED_BY As String = "Updated_By"

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

' --- QC_Number.csv AND QC_Number_Changes.csv Headers ---
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

' --- Label_Prints.csv Headers ---
Public Const BX_HEADER_QID_CODE As String = "QID_Code"
Public Const BX_HEADER_BATCH_NUMBER As String = "Batch_Number"
Public Const BX_HEADER_MANUF_LOT_NO As String = "Manuf_Lot_No"
Public Const BX_HEADER_MANUF_LOT_NO_2 As String = "Manuf_Lot_No_2"
Public Const BX_HEADER_EXPIRY As String = "Expiry"
Public Const BX_HEADER_QUANTITY As String = "Quantity"
Public Const BX_HEADER_CREATED_BY As String = "Created_By"
Public Const BX_HEADER_CREATED_DATE As String = "Created_date"
Public Const BX_HEADER_AVERY_LABEL_CODE As String = "Avery_Label_Code"
Public Const BX_HEADER_LABELS_PER_SHEET As String = "Labels_per_sheet"
Public Const BX_HEADER_TOTAL_PRINTED As String = "Total_Printed"
Public Const BX_HEADER_PRODUCT_ID As String = "Product_ID"
Public Const BX_HEADER_PMI_NUMBER As String = "CPMU_Ref"
Public Const BX_HEADER_PRODUCT_DESCRIPTION As String = "Product_Description"
Public Const BX_HEADER_PRINT_TYPE As String = "Print_Type"
Public Const BX_HEADER_REPRINT_REASON As String = "Reprint_Reason"
Public Const BX_HEADER_PRINTED As String = "Printed"
Public Const BX_HEADER_PRINT_DATE As String = "Print_Date"
Public Const BX_HEADER_PRINT_ID As String = "Print_ID"
Public Const BX_HEADER_ORIGINAL_PRINT_ID As String = "Original_Print_ID"
Public Const BX_HEADER_COMMENTS As String = "Comments"
Public Const BX_HEADER_REPRINT_REASON_TYPE As String = "Reprint_Reason_Type"
Public Const BX_HEADER_PRINTED_BY As String = "Printed_By"

' --- Raw_Material_Master.csv Headers ---
Public Const RMM_HEADER_RM_CODE As String = "RM_Code"
Public Const RMM_HEADER_STD_NAME As String = "Std_Name"
Public Const RMM_HEADER_ALT_NAME As String = "Alt_Name"
Public Const RMM_HEADER_CATEGORY As String = "Category"
Public Const RMM_HEADER_STORAGE_COND As String = "Storage_Cond"
Public Const RMM_HEADER_S4_NUMBER As String = "S4_Number"
Public Const RMM_HEADER_QUICK_CODE As String = "Quick_Code"
Public Const RMM_HEADER_PROD_ID As String = "Prod_ID"
Public Const RMM_HEADER_BRAND_NAME As String = "Brand_Name"
Public Const RMM_HEADER_PART_ARTG As String = "Part_ARTG"
Public Const RMM_HEADER_MASTER_UOM As String = "Master_UoM"
Public Const RMM_HEADER_QTY_PER_PACK As String = "QTY_Per_Pack"
Public Const RMM_HEADER_CONC_VOL As String = "Conc_Vol"
Public Const RMM_HEADER_WT_PER_ML As String = "Wt_per_mL"
Public Const RMM_HEADER_WT_PER_UNIT As String = "Wt_per_Unit"
Public Const RMM_HEADER_QC_NUMBER_REQUIRED As String = "QC_Number_Required"
Public Const RMM_HEADER_TESTING_FLAG As String = "Testing_Flag"
Public Const RMM_HEADER_QC_QTY As String = "QC_QTY"
Public Const RMM_HEADER_RETENTION_FLAG As String = "Retention_Flag"
Public Const RMM_HEADER_RETENTION_QTY As String = "Retention_QTY"
Public Const RMM_HEADER_SAMPLE_COMMENT As String = "Sample_Comment"
Public Const RMM_HEADER_CHEMALERT_SUBMIT As String = "ChemAlert_Submit"
Public Const RMM_HEADER_CHEMALERT_UPDATED As String = "ChemAlert_Updated"
Public Const RMM_HEADER_HAZCHEM As String = "HazChem"
Public Const RMM_HEADER_CAS_NUM As String = "CAS_Num"
Public Const RMM_HEADER_UN_CODE As String = "UN_Code"
Public Const RMM_HEADER_CURR_IND As String = "Curr_Ind"
Public Const RMM_HEADER_CREATED_BY As String = "Created_By"
Public Const RMM_HEADER_CREATE_DATE As String = "Create_Date"
Public Const RMM_HEADER_LAST_CHANGE_BY As String = "Last_Change_By"
Public Const RMM_HEADER_LAST_CHANGE_DATE As String = "Last_Change_Date"
Public Const RMM_HEADER_LAST_CHANGE_DETAIL As String = "Last_Change_Detail"
Public Const RMM_HEADER_EXTENDED_RETEST As String = "Extended_Retest"


' --- Physical_Inventory.csv Headers ---
Public Const PI_HEADER_DOC_NUMBER As String = "Doc_Number"
Public Const PI_HEADER_COUNTED_BY As String = "Counted_By"
Public Const PI_HEADER_COUNTED_DATE As String = "Counted_Date"
Public Const PI_HEADER_CODE As String = "Code"
Public Const PI_HEADER_DESCRIPTION As String = "Description"
Public Const PI_HEADER_BOOK_QTY As String = "Book_Qty"
Public Const PI_HEADER_COUNTED_QTY As String = "Counted_Qty"
Public Const PI_HEADER_DIFF_QTY As String = "Diff_Qty"
Public Const PI_HEADER_TRANS_ID As String = "RM_Trans_ID"

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
        Case Else ' AUTO or unknown -> try to detect DEV, then WORK, then HOME
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

' --- Utility: Array of All Label_Prints Headers ---
Public Function GetLabelPrintsHeaders() As Variant
    Dim headers(1 To 23) As String
    
    headers(1) = BX_HEADER_QID_CODE
    headers(2) = BX_HEADER_BATCH_NUMBER
    headers(3) = BX_HEADER_MANUF_LOT_NO
    headers(4) = BX_HEADER_MANUF_LOT_NO_2
    headers(5) = BX_HEADER_EXPIRY
    headers(6) = BX_HEADER_QUANTITY
    headers(7) = BX_HEADER_CREATED_BY
    headers(8) = BX_HEADER_CREATED_DATE
    headers(9) = BX_HEADER_AVERY_LABEL_CODE
    headers(10) = BX_HEADER_LABELS_PER_SHEET
    headers(11) = BX_HEADER_TOTAL_PRINTED
    headers(12) = BX_HEADER_PRODUCT_ID
    headers(13) = BX_HEADER_PMI_NUMBER
    headers(14) = BX_HEADER_PRODUCT_DESCRIPTION
    headers(15) = BX_HEADER_PRINT_TYPE
    headers(16) = BX_HEADER_REPRINT_REASON
    headers(17) = BX_HEADER_PRINTED
    headers(18) = BX_HEADER_PRINT_DATE
    headers(19) = BX_HEADER_PRINT_ID
    headers(20) = BX_HEADER_ORIGINAL_PRINT_ID
    headers(21) = BX_HEADER_COMMENTS
    headers(22) = BX_HEADER_REPRINT_REASON_TYPE
    headers(23) = BX_HEADER_PRINTED_BY
    
    GetLabelPrintsHeaders = headers
End Function

' --- Utility: Array of RM_Master headers (ordered) ---
Public Function GetRMMasterHeaders() As Variant
    Dim headers(1 To 33) As String

    headers(1) = RMM_HEADER_RM_CODE
    headers(2) = RMM_HEADER_STD_NAME
    headers(3) = RMM_HEADER_ALT_NAME
    headers(4) = RMM_HEADER_CATEGORY
    headers(5) = RMM_HEADER_STORAGE_COND
    headers(6) = RMM_HEADER_S4_NUMBER
    headers(7) = RMM_HEADER_QUICK_CODE
    headers(8) = RMM_HEADER_PROD_ID
    headers(9) = RMM_HEADER_BRAND_NAME
    headers(10) = RMM_HEADER_PART_ARTG
    headers(11) = RMM_HEADER_MASTER_UOM
    headers(12) = RMM_HEADER_QTY_PER_PACK
    headers(13) = RMM_HEADER_CONC_VOL
    headers(14) = RMM_HEADER_WT_PER_ML
    headers(15) = RMM_HEADER_WT_PER_UNIT
    headers(16) = RMM_HEADER_QC_NUMBER_REQUIRED
    headers(17) = RMM_HEADER_TESTING_FLAG
    headers(18) = RMM_HEADER_QC_QTY
    headers(19) = RMM_HEADER_RETENTION_FLAG
    headers(20) = RMM_HEADER_RETENTION_QTY
    headers(21) = RMM_HEADER_SAMPLE_COMMENT
    headers(22) = RMM_HEADER_CHEMALERT_SUBMIT
    headers(23) = RMM_HEADER_CHEMALERT_UPDATED
    headers(24) = RMM_HEADER_HAZCHEM
    headers(25) = RMM_HEADER_CAS_NUM
    headers(26) = RMM_HEADER_UN_CODE
    headers(27) = RMM_HEADER_CURR_IND
    headers(28) = RMM_HEADER_CREATED_BY
    headers(29) = RMM_HEADER_CREATE_DATE
    headers(30) = RMM_HEADER_LAST_CHANGE_BY
    headers(31) = RMM_HEADER_LAST_CHANGE_DATE
    headers(32) = RMM_HEADER_LAST_CHANGE_DETAIL
    headers(33) = RMM_HEADER_EXTENDED_RETEST

    GetRMMasterHeaders = headers
End Function

' --- Utility: Array of All Physical_Inventory Headers ---
Public Function GetPhysicalInventoryHeaders() As Variant
    Dim headers(1 To 9) As String
    
    headers(1) = PI_HEADER_DOC_NUMBER
    headers(2) = PI_HEADER_COUNTED_BY
    headers(3) = PI_HEADER_COUNTED_DATE
    headers(4) = PI_HEADER_CODE
    headers(5) = PI_HEADER_DESCRIPTION
    headers(6) = PI_HEADER_BOOK_QTY
    headers(7) = PI_HEADER_COUNTED_QTY
    headers(8) = PI_HEADER_DIFF_QTY
    headers(9) = PI_HEADER_TRANS_ID
    
    GetPhysicalInventoryHeaders = headers
End Function