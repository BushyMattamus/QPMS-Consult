Attribute VB_Name = "Constants"
Option Explicit

' Master Switches
Public Const MASTER_ENV_SWITCH As String = "WORK" ' Options: "WORK", "HOME"
Public Const PROD_DEV As String = "PROD" ' Change to "PROD" for production

' --- Folder Names ---
Public Const FOLDER_TABLES As String = "Tables"
Public Const FOLDER_MASTER_DATA As String = "Master Data"
Public Const FOLDER_MATERIAL_MANAGEMENT As String = "Material Management"

' --- Master Data Base Path (relative to workbook parent) ---
Public Const PATH_MASTER_DATA_TABLES As String = "Master Data\Tables"
Public Const PATH_MATERIAL_MANAGEMENT_TABLES As String = "Material Management\Tables"

' --- File Names ---
Public Const FILE_MASTER_DATA_FINISHED_GOODS As String = "MasterDataFinishedGoods.csv"
Public Const FILE_FINISHED_GOODS_BATCH_NUMBER_MASTER As String = "FinishedGoodsBatchNumberMasterFile.csv"
Public Const FILE_FINISHED_GOODS_BATCH_NUMBER_AUDIT As String = "FinishedGoodsBatchNumberAuditFile.csv"

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

' Finished Goods Batch Status Constants
Public Const FGB_STATUS_CREATED As String = "Created"
Public Const FGB_STATUS_PRINTED As String = "Printed"
Public Const FGB_STATUS_INPROGRESS As String = "InProgress"
Public Const FGB_STATUS_COMPLETE As String = "Complete"
Public Const FGB_STATUS_RELEASED As String = "Released"
Public Const FGB_STATUS_CANCELLED As String = "Cancelled"
Public Const FGB_STATUS_REJECTED As String = "Rejected"