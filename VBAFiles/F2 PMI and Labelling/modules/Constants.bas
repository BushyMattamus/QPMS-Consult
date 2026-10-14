Attribute VB_Name = "Constants"
Option Explicit

' Data folder locations and the PROD_DEV environment flag are auto-detected in PathResolver.
' Approved Documents Paths
Public Const APPROVED_PMI_WORD_PATH As String = "\Pharmacy\2.MFG\Documents\PMIs\2 Word Approved\"
Public Const APPROVED_PMI_PDF_PATH As String = "\Pharmacy\2.MFG\Documents\PMIs\4 PDF Approved\"
Public Const PRODUCT_APPROVED_LABELS_PATH As String = "\Pharmacy\2.MFG\Documents\Labels\2 Approved\"
' Archive file paths
Public Const PRODUCTION_ARCHIVE_PMI_PATH As String = "\Pharmacy\2.MFG\Documents\Archived Documents\Batch Original\"
Public Const PRODUCTION_ARCHIVE_COMPLETED_PMI_PATH As String = "\Pharmacy\2.MFG\Documents\Archived Documents\Batch Completed\"
Public Const ARCHIVE_PRODUCT_LABELS_PATH As String = "\Pharmacy\2.MFG\Documents\Archived Documents\Product Labels\"
Public Const ARCHIVE_PMI_LABELS_PATH As String = "\Pharmacy\2.MFG\Documents\Archived Documents\PMI Labels\"

' --- File and Folder Names ---
Public Const FOLDER_TABLES As String = "Tables"

' Files used by workbook-local Tables
' Files are UTF8 CSVs with BOM encoding
Public Const FILE_QC_NUMBER As String = "QC_Number.csv"  ' UTF-8 CSV with BOM
Public Const FILE_RM_TRANS As String = "RM_Trans.csv"  ' UTF-8 CSV without BOM
Public Const FILE_LABEL_DETAILS As String = "Label_Prints.csv" ' UTF-8 CSV without BOM
Public Const FILE_QC_NUMBER_CHANGES As String = "QC_Number_Changes.csv" ' UTF-8 CSV with BOM
Public Const FILE_RM_MASTER As String = "Raw_Material_Master.csv" ' UTF-8 CSV with BOM
Public Const FILE_PHYSICAL_INVENTORY As String = "Physical_Inventory.csv" ' UTF-8 CSV with BOM
Public Const FILE_FINISHED_GOODS_BATCH_MASTER As String = "FinishedGoodsBatchNumberMasterFile.csv" ' UTF-8 CSV without BOM
Public Const FILE_FINISHED_GOODS_BATCH_AUDIT As String = "FinishedGoodsBatchNumberAuditFile.csv" ' UTF-8 CSV without BOM
Public Const FILE_FG_MASTER_DATA As String = "MasterDataFinishedGoods.csv"
Public Const FILE_LABEL_REGISTER As String = "Label_Register.csv"

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

' --- FinishedGoodsBatchNumberMasterFile.csv Headers ---
Public Const FGBNM_HEADER_BATCH_NUMBER As String = "Batch_Number"
Public Const FGBNM_HEADER_PMI_NUMBER As String = "PMI_Number"
Public Const FGBNM_HEADER_PMI_VERSION As String = "PMI_Version"
Public Const FGBNM_HEADER_PRODUCT_DESCRIPTION As String = "Product_Description"
Public Const FGBNM_HEADER_SHELF_LIFE_NO As String = "Shelf_life_No"
Public Const FGBNM_HEADER_SHELF_LIFE As String = "Shelf_Life"
Public Const FGBNM_HEADER_MANUFACTURING_DATE As String = "Maunfacturing_Date"
Public Const FGBNM_HEADER_EXPIRY_DATE As String = "Expiry_Date"
Public Const FGBNM_HEADER_DATE_CREATED As String = "Date_Created"
Public Const FGBNM_HEADER_TIME_CREATED As String = "Time_Created"
Public Const FGBNM_HEADER_BATCH_STATUS As String = "Batch_Status"
Public Const FGBNM_HEADER_CREATED_BY As String = "Created_By"
Public Const FGBNM_HEADER_DATE_UPDATED As String = "Date_Updated"
Public Const FGBNM_HEADER_UPDATED_BY As String = "Updated_By"
Public Const FGBNM_HEADER_UPDATE_COMMENTS As String = "Update_Comments"
Public Const FGBNM_HEADER_SCANNED As String = "Scanned"


' --- finishedgoodsbatchnumberauditfile.csv Batch Status Values ---
' The Audit File has the same fields as the master file, with the additional field at the end
Public Const FGBNA_HEADER_TRANS_ID As String = "Trans_Id"

' PMI Printing Constants
Public Const TEXT_HOLDER_BATCH As String = "?BATCH"
Public Const TEXT_HOLDER_BARCODE As String = "?BXBARCODE"
Public Const TEXT_HOLDER_EXP As String = "?EXP"
Public Const TEXT_HOLDER_ISSUED_BY As String = "Batch Issued by:"

' Finished Goods Batch Status Constants
Public Const FGB_STATUS_CREATED As String = "Created"
Public Const FGB_STATUS_PRINTED As String = "Printed"
Public Const FGB_STATUS_INPROGRESS As String = "InProgress"
Public Const FGB_STATUS_COMPLETE As String = "Complete"
Public Const FGB_STATUS_RELEASED As String = "Released"
Public Const FGB_STATUS_CANCELLED As String = "Cancelled"
Public Const FGB_STATUS_REJECTED As String = "Rejected"

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
Public Const LABEL_HEADER_QID_CODE As String = "QID_Code"
Public Const LABEL_HEADER_BATCH_NUMBER As String = "Batch_Number"
Public Const LABEL_HEADER_MANUF_LOT_NO As String = "Mfg_Lot_No"
Public Const LABEL_HEADER_MANUF_LOT_NO_2 As String = "Mfg_Lot_No_2"
Public Const LABEL_HEADER_EXPIRY As String = "Expiry"
Public Const LABEL_HEADER_QUANTITY As String = "Quantity"
Public Const LABEL_HEADER_CREATED_BY As String = "Created_By"
Public Const LABEL_HEADER_CREATED_DATE As String = "Created_Date"
Public Const LABEL_HEADER_AVERY_LABEL_CODE As String = "Avery_Label_Code"
Public Const LABEL_HEADER_LABELS_PER_SHEET As String = "Labels_per_sheet"
Public Const LABEL_HEADER_TOTAL_PRINTED As String = "Total_Printed"
Public Const LABEL_HEADER_PRODUCT_ID As String = "Product_ID"
Public Const LABEL_HEADER_PMI_NUMBER As String = "CPMU_Ref"
Public Const LABEL_HEADER_PRODUCT_DESCRIPTION As String = "Product_Description"
Public Const LABEL_HEADER_PRINT_TYPE As String = "Print_Type"
Public Const LABEL_HEADER_REPRINT_REASON As String = "Reprint_Reason"
Public Const LABEL_HEADER_PRINTED As String = "Printed"
Public Const LABEL_HEADER_PRINT_DATE As String = "Print_Date"
Public Const LABEL_HEADER_PRINT_ID As String = "Print_ID"
Public Const LABEL_HEADER_ORIGINAL_PRINT_ID As String = "Original_Print_ID"
Public Const LABEL_HEADER_COMMENTS As String = "Comments"
Public Const LABEL_HEADER_REPRINT_REASON_TYPE As String = "Reprint_Reason_Type"
Public Const LABEL_HEADER_PRINTED_BY As String = "Printed_By"

' --- Label_Register.csv Headers ---
Public Const LR_HEADER_PMI_NUMBER As String = "PMI_Number"
Public Const LR_HEADER_AVERY_CODE As String = "Avery_Code"
Public Const LR_HEADER_UPDATED_BY As String = "Updated_By"
Public Const LR_HEADER_LAST_CHANGE_DATE As String = "Last_Change_Date"
Public Const LR_HEADER_SPECIAL_INSTRUCTIONS As String = "Special_Instructions"
Public Const LR_HEADER_LABELS_PER_BATCH As String = "Labels_Per_Batch"

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
    
    headers(1) = LABEL_HEADER_QID_CODE
    headers(2) = LABEL_HEADER_BATCH_NUMBER
    headers(3) = LABEL_HEADER_MANUF_LOT_NO
    headers(4) = LABEL_HEADER_MANUF_LOT_NO_2
    headers(5) = LABEL_HEADER_EXPIRY
    headers(6) = LABEL_HEADER_QUANTITY
    headers(7) = LABEL_HEADER_CREATED_BY
    headers(8) = LABEL_HEADER_CREATED_DATE
    headers(9) = LABEL_HEADER_AVERY_LABEL_CODE
    headers(10) = LABEL_HEADER_LABELS_PER_SHEET
    headers(11) = LABEL_HEADER_TOTAL_PRINTED
    headers(12) = LABEL_HEADER_PRODUCT_ID
    headers(13) = LABEL_HEADER_PMI_NUMBER
    headers(14) = LABEL_HEADER_PRODUCT_DESCRIPTION
    headers(15) = LABEL_HEADER_PRINT_TYPE
    headers(16) = LABEL_HEADER_REPRINT_REASON
    headers(17) = LABEL_HEADER_PRINTED
    headers(18) = LABEL_HEADER_PRINT_DATE
    headers(19) = LABEL_HEADER_PRINT_ID
    headers(20) = LABEL_HEADER_ORIGINAL_PRINT_ID
    headers(21) = LABEL_HEADER_COMMENTS
    headers(22) = LABEL_HEADER_REPRINT_REASON_TYPE
    headers(23) = LABEL_HEADER_PRINTED_BY
    
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

Public Function GetFinishedGoodsBatchMasterHeaders() As Variant
    Dim headers(1 To 16) As String
    
    headers(1) = FGBNM_HEADER_BATCH_NUMBER
    headers(2) = FGBNM_HEADER_PMI_NUMBER
    headers(3) = FGBNM_HEADER_PMI_VERSION
    headers(4) = FGBNM_HEADER_PRODUCT_DESCRIPTION
    headers(5) = FGBNM_HEADER_SHELF_LIFE_NO
    headers(6) = FGBNM_HEADER_SHELF_LIFE
    headers(7) = FGBNM_HEADER_MANUFACTURING_DATE
    headers(8) = FGBNM_HEADER_EXPIRY_DATE
    headers(9) = FGBNM_HEADER_DATE_CREATED
    headers(10) = FGBNM_HEADER_TIME_CREATED
    headers(11) = FGBNM_HEADER_BATCH_STATUS
    headers(12) = FGBNM_HEADER_CREATED_BY
    headers(13) = FGBNM_HEADER_DATE_UPDATED
    headers(14) = FGBNM_HEADER_UPDATED_BY
    headers(15) = FGBNM_HEADER_UPDATE_COMMENTS
    headers(16) = FGBNM_HEADER_SCANNED
    
    GetFinishedGoodsBatchMasterHeaders = headers
End Function

Public Function GetLabelRegisterHeaders() As Variant
    Dim headers(1 To 6) As String
    
    headers(1) = LR_HEADER_PMI_NUMBER
    headers(2) = LR_HEADER_AVERY_CODE
    headers(3) = LR_HEADER_UPDATED_BY
    headers(4) = LR_HEADER_LAST_CHANGE_DATE
    headers(5) = LR_HEADER_SPECIAL_INSTRUCTIONS
    headers(6) = LR_HEADER_LABELS_PER_BATCH
    
    GetLabelRegisterHeaders = headers
End Function