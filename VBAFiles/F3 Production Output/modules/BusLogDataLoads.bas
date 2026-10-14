Attribute VB_Name = "BusLogDataLoads"
Option Explicit

' ========================================================================
' FG Batch Details Import with Date Conversion
' ========================================================================
' Imports the Finished Goods Batch Master file into the "Batch_Details" sheet.
' Converts date columns (Manufacturing_Date, Expiry_Date, Date_Created) to dd/mm/yyyy format.
' ========================================================================

Public Sub FGBatchDetails()
    On Error GoTo ErrHandler

    Dim fgbdCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long

    ' --- Resolve source CSV path and destination sheet ---
    fgbdCSVPath = GetWorkbookTableFilePath(FILE_FINISHED_GOODS_BATCH_MASTER)
    Set wsDest = ThisWorkbook.Worksheets("Batch_Details")

    ' --- Clear existing contents and import CSV ---
    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(fgbdCSVPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 105, , ERROR_CSV_IMPORT & fgbdCSVPath
    End If

    ' --- Update named range NRpmiBatchNum to A2:A(lastRow) ---
    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ThisWorkbook.Names("NRpmiBatchNum").RefersTo = wsDest.Range("A2:A" & lastRow)
        ThisWorkbook.Names("NRpmiBatchStatus").RefersTo = wsDest.Range("K2:K" & lastRow)
        
        ' --- Populate Quick_Code in column P using FG_Master lookup ---
        wsDest.Cells(1, 16).value = "Quick_Code"
        PopulateQuickCodeFromMaster wsDest, lastRow
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "FG Batch Detail Load Error: " & Err.Description, vbCritical, "FGBatchDetails"
    GoTo CleanExit

CleanExit:
    Set wsDest = Nothing
    Exit Sub
End Sub

Private Sub PopulateQuickCodeFromMaster(ByVal wsBatchDetails As Worksheet, ByVal lastRow As Long)
    On Error GoTo ErrHandler

    Dim wsMaster As Worksheet
    Dim i As Long
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim lookupKey As String
    Dim masterRow As Long
    Dim quickCode As String

    Set wsMaster = ThisWorkbook.Worksheets("FG_Master")

    For i = 2 To lastRow
        pmiNumber = wsBatchDetails.Cells(i, 2).value
        pmiVersion = wsBatchDetails.Cells(i, 3).value
        
        If pmiNumber <> "" And pmiVersion <> "" Then
            lookupKey = pmiNumber & "|" & pmiVersion
            masterRow = FindMasterRow(wsMaster, pmiNumber, pmiVersion)
            
            If masterRow > 0 Then
                quickCode = wsMaster.Cells(masterRow, 5).value
                wsBatchDetails.Cells(i, 16).value = quickCode
            End If
        End If
    Next i

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PopulateQuickCodeFromMaster"
    GoTo CleanExit

CleanExit:
    Set wsMaster = Nothing
    Exit Sub
End Sub

Private Function FindMasterRow(ByVal wsMaster As Worksheet, ByVal pmiCode As String, ByVal pmiVersion As String) As Long
    Dim i As Long
    Dim masterLastRow As Long

    masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row

    For i = 2 To masterLastRow
        If wsMaster.Cells(i, 1).value = pmiCode And wsMaster.Cells(i, 2).value = pmiVersion Then
            FindMasterRow = i
            Exit Function
        End If
    Next i

    FindMasterRow = 0
End Function

Public Sub FGMasterDataLoad()
    On Error GoTo ErrHandler

    Dim fgMasterCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim sourceFolderPath As String

    sourceFolderPath = GetMasterDataTablesPath()
    fgMasterCSVPath = FILE_FG_MASTER_DATA
    Set wsDest = ThisWorkbook.Worksheets("FG_Master")

    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(fgMasterCSVPath, wsDest, wsDest.Range("A1"), True, sourceFolderPath)
    If Not importSuccess Then
        Err.Raise vbObjectError + 106, , ERROR_CSV_IMPORT & fgMasterCSVPath
    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "FG Master Data Load Error: " & Err.Description, vbCritical, "FGMasterDataLoad"
    GoTo CleanExit

CleanExit:
    Set wsDest = Nothing
    Exit Sub
End Sub