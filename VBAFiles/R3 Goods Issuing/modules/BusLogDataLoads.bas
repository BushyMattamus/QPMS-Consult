Attribute VB_Name = "BusLogDataLoads"
Option Explicit

' Imports specified columns from QC_Number.csv into the "QC_Number" sheet
' Uses QueryTable (TEXT connection) to avoid opening the CSV as a workbook
' Reuses public helpers: FindHeaderCol, RM_Pad4, MakeTempSheetName, SheetExists

Public Sub LoadQCNumberData()
    Dim qcCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long
    Dim r As Long
    Dim s As String

    On Error GoTo ErrorHandler
    
    ' Resolve source CSV path and destination sheet
    qcCSVPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    Set wsDest = ThisWorkbook.Worksheets("QC_Number")

    ' Clear existing contents and import whole CSV into the sheet
    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(qcCSVPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 105, , ERROR_CSV_IMPORT & qcCSVPath
    End If

    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 2 Then
        wsDest.Columns(1).NumberFormat = "@"
        For r = 2 To lastRow
            s = CStr(wsDest.Cells(r, 1).value)
            s = Trim$(s)
            If Len(s) = 0 Then
                wsDest.Cells(r, 1).value = ""
            ElseIf Len(s) >= 6 Then
                wsDest.Cells(r, 1).value = s
            Else
                wsDest.Cells(r, 1).value = Right$("000000" & s, 6)
            End If
        Next r
    End If

    wsDest.Columns(2).NumberFormat = "@"
    If lastRow >= 2 Then
        For r = 2 To lastRow
            s = CStr(wsDest.Cells(r, 2).value)
            s = Trim$(s)
            If Len(s) = 0 Then
                wsDest.Cells(r, 2).value = ""
            ElseIf Len(s) >= 4 Then
                wsDest.Cells(r, 2).value = s
            Else
                wsDest.Cells(r, 2).value = Right$("0000" & s, 4)
            End If
        Next r
        
        ThisWorkbook.Names("NRqcNumBalance").RefersTo = wsDest.Range("U2:U" & lastRow)
        ThisWorkbook.Names("NRqcNumFirstRelease").RefersTo = wsDest.Range("W2:W" & lastRow)
        ThisWorkbook.Names("NRqcNumLastRetest").RefersTo = wsDest.Range("AC2:AC" & lastRow)
        ThisWorkbook.Names("NRqcNumMfgLotNumber").RefersTo = wsDest.Range("I2:I" & lastRow)
        ThisWorkbook.Names("NRqcNumMfgName").RefersTo = wsDest.Range("H2:H" & lastRow)
        ThisWorkbook.Names("NRqcNumNilStockFlag").RefersTo = wsDest.Range("AE2:AE" & lastRow)
        ThisWorkbook.Names("NRqcNumNumber").RefersTo = wsDest.Range("A2:A" & lastRow)
        ThisWorkbook.Names("NRqcNumRMCode").RefersTo = wsDest.Range("B2:B" & lastRow)
        ThisWorkbook.Names("NRqcNumRMDesc").RefersTo = wsDest.Range("C2:C" & lastRow)
        
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "LoadQCNumberData error: " & Err.Description, vbCritical
    Exit Sub

End Sub

Public Sub LoadRMMasterData()
    ' Sheet: RM_Master (output)
    ' Imports Raw_Material_Master.csv (as TEXT via QueryTable to a temp sheet)
    ' and writes selected columns to RM_Master.
    
    Dim conn As Object
    Dim rs As Object
    Dim rmCSVPath As String
    Dim fileName As String
    Dim sql As String
    Dim wsDest As Worksheet
    Dim wsStart As Worksheet
    Dim outRow As Long

    On Error GoTo ErrorHandler

    ' --- Remember starting sheet ---
    Set wsStart = ActiveSheet

    ' --- Fixed source path using constants ---
    rmCSVPath = GetWorkbookTableFilePath(FILE_RM_MASTER)

    ' --- Check if file exists ---
    If Not FileExists(rmCSVPath) Then
        Err.Raise vbObjectError + 102, , "File not found: " & rmCSVPath
    End If

    ' --- Destination sheet ---
    Set wsDest = ThisWorkbook.Worksheets("RM_Master")
    wsDest.Cells.ClearContents

    ' --- Open ADO connection ---
    Set conn = OpenADOConnection(Left(rmCSVPath, InStrRev(rmCSVPath, "\") - 1))
    If conn Is Nothing Then
        Err.Raise vbObjectError + 103, , "Failed to open ADO connection."
    End If

    ' --- SQL query to get all columns (avoids BOM issue) ---
    fileName = Mid(rmCSVPath, InStrRev(rmCSVPath, "\") + 1)
    sql = "SELECT * FROM [" & fileName & "]"

    ' --- Execute query ---
    Set rs = conn.Execute(sql)

    ' --- Write output headers ---
    wsDest.Range("A1:Q1").value = Array( _
        RMM_HEADER_RM_CODE, RMM_HEADER_STD_NAME, RMM_HEADER_ALT_NAME, RMM_HEADER_MASTER_UOM, RMM_HEADER_WT_PER_ML, RMM_HEADER_CURR_IND, _
        RMM_HEADER_TESTING_FLAG, RMM_HEADER_QC_QTY, RMM_HEADER_RETENTION_FLAG, RMM_HEADER_RETENTION_QTY, RMM_HEADER_STORAGE_COND, _
        RMM_HEADER_QUICK_CODE, RMM_HEADER_BRAND_NAME, RMM_HEADER_PART_ARTG, RMM_HEADER_QC_NUMBER_REQUIRED, RMM_HEADER_SAMPLE_COMMENT, _
        RMM_HEADER_EXTENDED_RETEST)

    ' --- Load data into memory array ---
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim colCount As Long
    
    colCount = 17
    rowCount = 0
    
    ' --- First pass: count rows ---
    Do While Not rs.EOF
        rowCount = rowCount + 1
        rs.MoveNext
    Loop
    
    ' --- Check if we have data ---
    If rowCount = 0 Then
        GoTo CleanExit
    End If
    
    ' --- Size the array ---
    ReDim dataArray(1 To rowCount, 1 To colCount)
    
    ' --- Reset recordset to beginning ---
    rs.MoveFirst
    
    ' --- Second pass: populate array ---
    outRow = 1
    Do While Not rs.EOF
        dataArray(outRow, 1) = RM_Pad4(CStr(GetFieldValue(rs, RMM_HEADER_RM_CODE) & ""))
        dataArray(outRow, 2) = GetFieldValue(rs, RMM_HEADER_STD_NAME) & ""
        dataArray(outRow, 3) = GetFieldValue(rs, RMM_HEADER_ALT_NAME) & ""
        dataArray(outRow, 4) = GetFieldValue(rs, RMM_HEADER_MASTER_UOM) & ""
        dataArray(outRow, 5) = GetFieldValue(rs, RMM_HEADER_WT_PER_ML) & ""
        dataArray(outRow, 6) = GetFieldValue(rs, RMM_HEADER_CURR_IND) & ""
        dataArray(outRow, 7) = GetFieldValue(rs, RMM_HEADER_TESTING_FLAG) & ""
        dataArray(outRow, 8) = GetFieldValue(rs, RMM_HEADER_QC_QTY) & ""
        dataArray(outRow, 9) = GetFieldValue(rs, RMM_HEADER_RETENTION_FLAG) & ""
        dataArray(outRow, 10) = GetFieldValue(rs, RMM_HEADER_RETENTION_QTY) & ""
        dataArray(outRow, 11) = GetFieldValue(rs, RMM_HEADER_STORAGE_COND) & ""
        dataArray(outRow, 12) = GetFieldValue(rs, RMM_HEADER_QUICK_CODE) & ""
        dataArray(outRow, 13) = GetFieldValue(rs, RMM_HEADER_BRAND_NAME) & ""
        dataArray(outRow, 14) = GetFieldValue(rs, RMM_HEADER_PART_ARTG) & ""
        dataArray(outRow, 15) = GetFieldValue(rs, RMM_HEADER_QC_NUMBER_REQUIRED) & ""
        dataArray(outRow, 16) = GetFieldValue(rs, RMM_HEADER_SAMPLE_COMMENT) & ""
        dataArray(outRow, 17) = GetFieldValue(rs, RMM_HEADER_EXTENDED_RETEST) & ""
        outRow = outRow + 1
        rs.MoveNext
    Loop
    
    ' --- Format column A as text before writing ---
    wsDest.Columns(1).NumberFormat = "@"
    
    ' --- Write entire array to worksheet in one operation ---
    wsDest.Range("A2").Resize(rowCount, colCount).value = dataArray

    GoTo CleanExit

ErrorHandler:
    MsgBox "LoadRMMasterData error: " & Err.Description, vbCritical, "LoadRMMasterData"

CleanExit:
    If Not rs Is Nothing Then
        If Not rs.State = 0 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If Not conn.State = 0 Then conn.Close
        Set conn = Nothing
    End If
    If Not wsStart Is Nothing Then wsStart.Activate
    Set wsStart = Nothing
    Set wsDest = Nothing
    
End Sub

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

    ' --- Update named range NRbatchNum to A2:A(lastRow) ---
    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ThisWorkbook.Names("NRbatchNum").RefersTo = wsDest.Range("A2:A" & lastRow)
        ' ThisWorkbook.Names("NRbatchPMINum").RefersTo = wsDest.Range("B2:K" & lastRow)
        ' ThisWorkbook.Names("NRbatchProdDesc").RefersTo = wsDest.Range("D2:D" & lastRow)
        ThisWorkbook.Names("NRbatchStatus").RefersTo = wsDest.Range("K2:K" & lastRow)
        
        ' --- Populate Quick_Code in column P using FG_Master lookup ---
        wsDest.Cells(1, 16).value = "Quick_Code"
        PopulateQuickCodeFromMaster wsDest, lastRow
        
        ' ThisWorkbook.Names("NRbatchQuickCode").RefersTo = wsDest.Range("P2:P" & lastRow)
        
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
    Dim masterLastRow As Long
    Dim masterData As Variant
    Dim batchData As Variant
    Dim quickOut() As Variant
    Dim quickCodes As Object
    Dim i As Long
    Dim pmiNumber As String
    Dim pmiVersion As String
    Dim lookupKey As String

    If lastRow < 2 Then GoTo CleanExit

    Set wsMaster = ThisWorkbook.Worksheets("FG_Master")
    masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    If masterLastRow < 2 Then GoTo CleanExit

    ' Read master once: columns A (PMI), B (version), E (quick code). First match wins.
    masterData = wsMaster.Range(wsMaster.Cells(2, 1), wsMaster.Cells(masterLastRow, 5)).value
    Set quickCodes = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(masterData, 1)
        lookupKey = CStr(masterData(i, 1)) & "|" & CStr(masterData(i, 2))
        If Not quickCodes.Exists(lookupKey) Then quickCodes.Add lookupKey, masterData(i, 5)
    Next i

    ' Read batch rows once: columns B to P. Column P (index 15) keeps its value when there is no match.
    batchData = wsBatchDetails.Range(wsBatchDetails.Cells(2, 2), wsBatchDetails.Cells(lastRow, 16)).value
    ReDim quickOut(1 To lastRow - 1, 1 To 1)
    For i = 1 To lastRow - 1
        quickOut(i, 1) = batchData(i, 15)
        pmiNumber = CStr(batchData(i, 1))
        pmiVersion = CStr(batchData(i, 2))
        If pmiNumber <> "" And pmiVersion <> "" Then
            lookupKey = pmiNumber & "|" & pmiVersion
            If quickCodes.Exists(lookupKey) Then quickOut(i, 1) = quickCodes(lookupKey)
        End If
    Next i
    wsBatchDetails.Cells(2, 16).Resize(lastRow - 1, 1).value = quickOut

    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PopulateQuickCodeFromMaster"
    GoTo CleanExit

CleanExit:
    Set wsMaster = Nothing
    Set quickCodes = Nothing
    Exit Sub
End Sub


Public Sub FGMasterDataLoad()
    On Error GoTo ErrHandler

    Dim fgMasterCSVPath As String
    Dim wsDest As Worksheet
    Dim importSuccess As Boolean
    Dim lastRow As Long
    Dim sourceFolderPath As String

    sourceFolderPath = GetMasterDataTablesPath()
    fgMasterCSVPath = FILE_FG_MASTER_DATA
    Set wsDest = ThisWorkbook.Worksheets("FG_Master")

    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(fgMasterCSVPath, wsDest, wsDest.Range("A1"), True, sourceFolderPath)
    If Not importSuccess Then
        Err.Raise vbObjectError + 106, , ERROR_CSV_IMPORT & fgMasterCSVPath
    End If

'    ' --- Update named ranges ---
'    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
'    If lastRow >= 2 Then
'        ThisWorkbook.Names("NRpmi_CodeMaster").RefersTo = wsDest.Range("A2:A" & lastRow)
'        ThisWorkbook.Names("NRpmi_StatusMaster").RefersTo = wsDest.Range("G2:G" & lastRow)
'    End If

    GoTo CleanExit

ErrHandler:
    MsgBox "FG Master Data Load Error: " & Err.Description, vbCritical, "FGMasterDataLoad"
    GoTo CleanExit

CleanExit:
    Set wsDest = Nothing
    Exit Sub
End Sub
