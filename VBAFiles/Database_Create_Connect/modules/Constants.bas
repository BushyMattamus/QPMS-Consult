Attribute VB_Name = "Constants"
'=============================================================================
' Module: Constants
' Purpose: Define database paths, table schemas, and field definitions
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' DATABASE PATHS
'------------------------------------------------------------------------------
Public Const DB_PATH As String = "DataFiles\Database.accdb"
Public Const CSV_DATA_DIR As String = "DataFiles\Master Data\Tables\"

'------------------------------------------------------------------------------
' TABLE NAMES (matching CSV files)
'------------------------------------------------------------------------------
Public Const TABLE_MASTER_FINISHED_GOODS As String = "MasterDataFinishedGoods"
Public Const TABLE_RAW_MATERIAL_MASTER As String = "RawMaterialMaster"
Public Const TABLE_AUDIT_TRAIL As String = "AuditTrail"
Public Const TABLE_RM_CHANGES As String = "RMChanges"

'------------------------------------------------------------------------------
' FIELD DEFINITIONS - MasterDataFinishedGoods (18 columns)
'------------------------------------------------------------------------------
Public Const FG_FIELD_ID As String = "ID"
Public Const FG_FIELD_PMI_CODE As String = "PMI_Code"
Public Const FG_FIELD_PMI_VERSION As String = "PMI_Version"
Public Const FG_FIELD_PRODUCT_TYPE As String = "Product_Type"
Public Const FG_FIELD_PRODUCT_CATEGORY As String = "Product_Category"
Public Const FG_FIELD_QUICK_CODE As String = "Quick_Code"
Public Const FG_FIELD_PRODUCT_DESCRIPTION As String = "Product_Description"
Public Const FG_FIELD_PRODUCT_STATUS As String = "Product_Status"
Public Const FG_FIELD_SHELF_LIFE As String = "Shelf_Life"
Public Const FG_FIELD_SHELF_LIFE_UNITS As String = "Shelf_life(No of days/Months/Years)"
Public Const FG_FIELD_PRODUCTION_TYPE As String = "Production_Type"
Public Const FG_FIELD_DATE_CREATED As String = "Date_Created"
Public Const FG_FIELD_CREATED_BY As String = "Created_By"
Public Const FG_FIELD_DATE_UPDATED As String = "Date_Updated"
Public Const FG_FIELD_UPDATED_BY As String = "Updated_By"
Public Const FG_FIELD_UPDATE_COMMENTS As String = "Update_Comments"
Public Const FG_FIELD_PMI_ACRONYM As String = "PMI_Acronym"
Public Const FG_FIELD_PMI_NUMBER As String = "PMI_Number"
Public Const FG_FIELD_OLD_PMI As String = "Old_PMI"
Public Const FG_FIELD_PROD_ID As String = "Prod_ID"

'------------------------------------------------------------------------------
' FIELD DEFINITIONS - RawMaterialMaster (36 columns)
'------------------------------------------------------------------------------
Public Const RM_FIELD_ID As String = "ID"
Public Const RM_FIELD_RM_CODE As String = "RM_Code"
Public Const RM_FIELD_STD_NAME As String = "Std_Name"
Public Const RM_FIELD_ALT_NAME As String = "Alt_Name"
Public Const RM_FIELD_CATEGORY As String = "Category"
Public Const RM_FIELD_STORAGE_COND As String = "Storage_Cond"
Public Const RM_FIELD_S4_NUMBER As String = "S4_Number"
Public Const RM_FIELD_QUICK_CODE As String = "Quick_Code"
Public Const RM_FIELD_PROD_ID As String = "Prod_ID"
Public Const RM_FIELD_BRAND_NAME As String = "Brand_Name"
Public Const RM_FIELD_PART_ARTG As String = "Part_ARTG"
Public Const RM_FIELD_MASTER_UOM As String = "Master_UoM"
Public Const RM_FIELD_QTY_PER_PACK As String = "QTY_Per_Pack"
Public Const RM_FIELD_CONC_VOL As String = "Conc_Vol"
Public Const RM_FIELD_WT_PER_ML As String = "Wt_per_mL"
Public Const RM_FIELD_WT_PER_UNIT As String = "Wt_per_Unit"
Public Const RM_FIELD_QC_NUMBER_REQUIRED As String = "QC_Number_Required"
Public Const RM_FIELD_TESTING_FLAG As String = "Testing_Flag"
Public Const RM_FIELD_QC_QTY As String = "QC_QTY"
Public Const RM_FIELD_RETENTION_FLAG As String = "Retention_Flag"
Public Const RM_FIELD_RETENTION_QTY As String = "Retention_QTY"
Public Const RM_FIELD_SAMPLE_COMMENT As String = "Sample_Comment"
Public Const RM_FIELD_CHEM_ALERT_SUBMIT As String = "ChemAlert_Submit"
Public Const RM_FIELD_CHEM_ALERT_UPDATED As String = "ChemAlert_Updated"
Public Const RM_FIELD_HAZCHEM As String = "HazChem"
Public Const RM_FIELD_CAS_NUM As String = "CAS_Num"
Public Const RM_FIELD_UN_CODE As String = "UN_Code"
Public Const RM_FIELD_CURR_IND As String = "Curr_Ind"
Public Const RM_FIELD_CREATED_BY As String = "Created_By"
Public Const RM_FIELD_CREATE_DATE As String = "Create_Date"
Public Const RM_FIELD_LAST_CHANGE_BY As String = "Last_Change_By"
Public Const RM_FIELD_LAST_CHANGE_DATE As String = "Last_Change_Date"
Public Const RM_FIELD_LAST_CHANGE_DETAIL As String = "Last_Change_Detail"
Public Const RM_FIELD_EXTENDED_RETEST As String = "Extended_Retest"
Public Const RM_FIELD_MRP_MIN As String = "MRP_Min"
Public Const RM_FIELD_MRP_MAX As String = "MRP_Max"
Public Const RM_FIELD_PACK_PRICE As String = "Pack_Price"

'------------------------------------------------------------------------------
' FIELD DEFINITIONS - AuditTrail (18 columns)
'------------------------------------------------------------------------------
Public Const AUDIT_FIELD_ID As String = "ID"
Public Const AUDIT_FIELD_PMI_CODE As String = "PMI_Code"
Public Const AUDIT_FIELD_PMI_VERSION As String = "PMI_Version"
Public Const AUDIT_FIELD_PRODUCT_TYPE As String = "Product_Type"
Public Const AUDIT_FIELD_PRODUCT_CATEGORY As String = "Product_Category"
Public Const AUDIT_FIELD_QUICK_CODE As String = "Quick_Code"
Public Const AUDIT_FIELD_PRODUCT_DESCRIPTION As String = "Product_Description"
Public Const AUDIT_FIELD_PRODUCT_STATUS As String = "Product_Status"
Public Const AUDIT_FIELD_SHELF_LIFE As String = "Shelf_Life"
Public Const AUDIT_FIELD_SHELF_LIFE_UNITS As String = "Shelf_life(No of days/Months/Years)"
Public Const AUDIT_FIELD_PRODUCTION_TYPE As String = "Production_Type"
Public Const AUDIT_FIELD_DATE_CREATED As String = "Date_Created"
Public Const AUDIT_FIELD_CREATED_BY As String = "Created_By"
Public Const AUDIT_FIELD_DATE_UPDATED As String = "Date_Updated"
Public Const AUDIT_FIELD_UPDATED_BY As String = "Updated_By"
Public Const AUDIT_FIELD_UPDATE_COMMENTS As String = "Update_Comments"
Public Const AUDIT_FIELD_PMI_ACRONYM As String = "PMI_Acronym"
Public Const AUDIT_FIELD_PMI_NUMBER As String = "PMI_Number"
Public Const AUDIT_FIELD_OLD_PMI As String = "Old_PMI"
Public Const AUDIT_FIELD_PROD_ID As String = "Prod_ID"

'------------------------------------------------------------------------------
' FIELD DEFINITIONS - RMChanges (36 columns, same as RawMaterialMaster)
'------------------------------------------------------------------------------
Public Const RMCHANGES_FIELD_ID As String = "ID"
Public Const RMCHANGES_FIELD_RM_CODE As String = "RM_Code"
Public Const RMCHANGES_FIELD_STD_NAME As String = "Std_Name"
Public Const RMCHANGES_FIELD_ALT_NAME As String = "Alt_Name"
Public Const RMCHANGES_FIELD_CATEGORY As String = "Category"
Public Const RMCHANGES_FIELD_STORAGE_COND As String = "Storage_Cond"
Public Const RMCHANGES_FIELD_S4_NUMBER As String = "S4_Number"
Public Const RMCHANGES_FIELD_QUICK_CODE As String = "Quick_Code"
Public Const RMCHANGES_FIELD_PROD_ID As String = "Prod_ID"
Public Const RMCHANGES_FIELD_BRAND_NAME As String = "Brand_Name"
Public Const RMCHANGES_FIELD_PART_ARTG As String = "Part_ARTG"
Public Const RMCHANGES_FIELD_MASTER_UOM As String = "Master_UoM"
Public Const RMCHANGES_FIELD_QTY_PER_PACK As String = "QTY_Per_Pack"
Public Const RMCHANGES_FIELD_CONC_VOL As String = "Conc_Vol"
Public Const RMCHANGES_FIELD_WT_PER_ML As String = "Wt_per_mL"
Public Const RMCHANGES_FIELD_WT_PER_UNIT As String = "Wt_per_Unit"
Public Const RMCHANGES_FIELD_QC_NUMBER_REQUIRED As String = "QC_Number_Required"
Public Const RMCHANGES_FIELD_TESTING_FLAG As String = "Testing_Flag"
Public Const RMCHANGES_FIELD_QC_QTY As String = "QC_QTY"
Public Const RMCHANGES_FIELD_RETENTION_FLAG As String = "Retention_Flag"
Public Const RMCHANGES_FIELD_RETENTION_QTY As String = "Retention_QTY"
Public Const RMCHANGES_FIELD_SAMPLE_COMMENT As String = "Sample_Comment"
Public Const RMCHANGES_FIELD_CHEM_ALERT_SUBMIT As String = "ChemAlert_Submit"
Public Const RMCHANGES_FIELD_CHEM_ALERT_UPDATED As String = "ChemAlert_Updated"
Public Const RMCHANGES_FIELD_HAZCHEM As String = "HazChem"
Public Const RMCHANGES_FIELD_CAS_NUM As String = "CAS_Num"
Public Const RMCHANGES_FIELD_UN_CODE As String = "UN_Code"
Public Const RMCHANGES_FIELD_CURR_IND As String = "Curr_Ind"
Public Const RMCHANGES_FIELD_CREATED_BY As String = "Created_By"
Public Const RMCHANGES_FIELD_CREATE_DATE As String = "Create_Date"
Public Const RMCHANGES_FIELD_LAST_CHANGE_BY As String = "Last_Change_By"
Public Const RMCHANGES_FIELD_LAST_CHANGE_DATE As String = "Last_Change_Date"
Public Const RMCHANGES_FIELD_LAST_CHANGE_DETAIL As String = "Last_Change_Detail"
Public Const RMCHANGES_FIELD_EXTENDED_RETEST As String = "Extended_Retest"
Public Const RMCHANGES_FIELD_MRP_MIN As String = "MRP_Min"
Public Const RMCHANGES_FIELD_MRP_MAX As String = "MRP_Max"
Public Const RMCHANGES_FIELD_PACK_PRICE As String = "Pack_Price"

'------------------------------------------------------------------------------
' DATA TYPE MAPPINGS (for Access table creation)
'------------------------------------------------------------------------------
' Text fields (up to 255 characters)
Public Const DT_TEXT_50 As String = "Text(50)"
Public Const DT_TEXT_100 As String = "Text(100)"
Public Const DT_TEXT_255 As String = "Text(255)"

' Number fields
Public Const DT_NUMBER As String = "Number"
Public Const DT_CURRENCY As String = "Currency"

' Date/Time fields
Public Const DT_DATE As String = "Date/Time"

'------------------------------------------------------------------------------
' PRIMARY KEY FIELD (AutoNumber)
'------------------------------------------------------------------------------
Public Const PK_FIELD As String = "ID"

'------------------------------------------------------------------------------
' ERROR CODES
'------------------------------------------------------------------------------
Public Const ERR_DB_NOT_CREATED As Long = 1001
Public Const ERR_TABLE_NOT_CREATED As Long = 1002
Public Const ERR_RELATIONSHIP_FAILED As Long = 1003
Public Const ERR_IMPORT_FAILED As Long = 1004
Public Const ERR_CONNECTION_FAILED As Long = 1005

'------------------------------------------------------------------------------
' STATUS CONSTANTS
'------------------------------------------------------------------------------
Public Const STATUS_SUCCESS As String = "Success"
Public Const STATUS_ERROR As String = "Error"
Public Const STATUS_PENDING As String = "Pending"
