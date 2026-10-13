Attribute VB_Name = "BusLogDataLoads"
Option Explicit

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
    rmCSVPath = GetMasterDataPath(FOLDER_TABLES & "\" & FILE_RM_MASTER)
    Debug.Print rmCSVPath

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
    wsDest.Range("A1:R1").value = Array( _
        HEADER_RM_CODE, HEADER_STD_NAME, HEADER_ALT_NAME, HEADER_MASTER_UOM, HEADER_WT_PER_ML, HEADER_CURR_IND, _
        HEADER_TESTING_FLAG, HEADER_QC_QTY, HEADER_RETENTION_FLAG, HEADER_RETENTION_QTY, HEADER_STORAGE_COND, _
        HEADER_QUICK_CODE, HEADER_BRAND_NAME, HEADER_PART_ARTG, HEADER_QC_NUMBER_REQUIRED, HEADER_SAMPLE_COMMENT, _
        HEADER_EXTENDED_RETEST_RM, HEADER_PACK_PRICE, HEADER_PO_NUMBER)

    ' --- Load data into memory array ---
    Dim dataArray() As Variant
    Dim rowCount As Long
    Dim colCount As Long
    
    colCount = 19
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
        dataArray(outRow, 1) = RM_Pad4(CStr(GetFieldValue(rs, HEADER_RM_CODE) & ""))
        dataArray(outRow, 2) = GetFieldValue(rs, HEADER_STD_NAME) & ""
        dataArray(outRow, 3) = GetFieldValue(rs, HEADER_ALT_NAME) & ""
        dataArray(outRow, 4) = GetFieldValue(rs, HEADER_MASTER_UOM) & ""
        dataArray(outRow, 5) = GetFieldValue(rs, HEADER_WT_PER_ML) & ""
        dataArray(outRow, 6) = GetFieldValue(rs, HEADER_CURR_IND) & ""
        dataArray(outRow, 7) = GetFieldValue(rs, HEADER_TESTING_FLAG) & ""
        dataArray(outRow, 8) = GetFieldValue(rs, HEADER_QC_QTY) & ""
        dataArray(outRow, 9) = GetFieldValue(rs, HEADER_RETENTION_FLAG) & ""
        dataArray(outRow, 10) = GetFieldValue(rs, HEADER_RETENTION_QTY) & ""
        dataArray(outRow, 11) = GetFieldValue(rs, HEADER_STORAGE_COND) & ""
        dataArray(outRow, 12) = GetFieldValue(rs, HEADER_QUICK_CODE) & ""
        dataArray(outRow, 13) = GetFieldValue(rs, HEADER_BRAND_NAME) & ""
        dataArray(outRow, 14) = GetFieldValue(rs, HEADER_PART_ARTG) & ""
        dataArray(outRow, 15) = GetFieldValue(rs, HEADER_QC_NUMBER_REQUIRED) & ""
        dataArray(outRow, 16) = GetFieldValue(rs, HEADER_SAMPLE_COMMENT) & ""
        dataArray(outRow, 17) = GetFieldValue(rs, HEADER_EXTENDED_RETEST_RM) & ""
        dataArray(outRow, 18) = GetFieldValue(rs, HEADER_PACK_PRICE) & ""
        dataArray(outRow, 19) = GetFieldValue(rs, HEADER_PO_NUMBER) & ""
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


' Imports Vendor_List.csv into the "Vendor_List" sheet
' Uses QueryTable (TEXT connection) to avoid opening the CSV as a workbook
' Reuses public helpers: FindHeaderCol, MakeTempSheetName, SheetExists

Public Sub LoadVendorList()

    Dim vendorCSVPath As String
    Dim wsDest As Worksheet
    Dim wsMaintain As Worksheet
    Dim wsStart As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long
    Dim importSuccess As Boolean

    On Error GoTo ErrorHandler
    
    ' --- Remember starting sheet ---
    Set wsStart = ActiveSheet

    vendorCSVPath = GetWorkbookTableFilePath(FILE_VENDOR_LIST)

    ' --- Destination sheets ---
    Set wsDest = ThisWorkbook.Worksheets("Vendor_List")
    Set wsMaintain = ThisWorkbook.Worksheets("Maintain Vendor")
    wsDest.Cells.ClearContents
    wsMaintain.Range("B4:B10000").ClearContents

    ActiveWorkbook.Protect Structure:=False, Windows:=False

    importSuccess = ImportCSVWithOpenText(vendorCSVPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 104, , ERROR_CSV_IMPORT & vendorCSVPath
    End If

    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row

    outRow = 2
    For r = 2 To lastRow
        wsMaintain.Cells(outRow + 3, 2).NumberFormat = "@"
        wsMaintain.Cells(outRow + 3, 2).value = wsDest.Cells(r, 1).value
        outRow = outRow + 1
    Next r

Cleanup:
    On Error Resume Next
    wsDest.Visible = xlSheetHidden
    ActiveWorkbook.Protect Structure:=True, Windows:=True
    
    If Not wsStart Is Nothing Then wsStart.Activate
    
    Set wsStart = Nothing
    Set wsDest = Nothing
    Set wsMaintain = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "LoadVendorList error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub


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
    
    AppStateUtils.DisableAppUpdates

    qcCSVPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    Set wsDest = ThisWorkbook.Worksheets("QC_Number")

    wsDest.Cells.ClearContents
    importSuccess = ImportCSVWithOpenText(qcCSVPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 105, , ERROR_CSV_IMPORT & qcCSVPath
    End If

    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    
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
    End If
    
    AppStateUtils.RestoreAppState

    Exit Sub
    
ErrorHandler:
    MsgBox "LoadQCNumberData error: " & Err.Description, vbCritical
    AppStateUtils.RestoreAppState
    Exit Sub

End Sub

Public Sub LoadManufacturerList()

    Dim manuCSVPath As String
    Dim wsDest As Worksheet
    Dim wsMaintain As Worksheet
    Dim wsStart As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long
    Dim importSuccess As Boolean

    On Error GoTo ErrorHandler
    
    Set wsStart = ActiveSheet

    manuCSVPath = GetWorkbookTableFilePath(FILE_MANUFACTURER_LIST)

    Set wsDest = ThisWorkbook.Worksheets("Manufacturer_List")
    Set wsMaintain = ThisWorkbook.Worksheets("Maintain Manufacturer")
    wsDest.Cells.ClearContents
    wsMaintain.Range("B4:B10000").ClearContents

    ActiveWorkbook.Protect Structure:=False, Windows:=False

    importSuccess = ImportCSVWithOpenText(manuCSVPath, wsDest, wsDest.Range("A1"), True)
    If Not importSuccess Then
        Err.Raise vbObjectError + 104, , ERROR_CSV_IMPORT & manuCSVPath
    End If

    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row

    outRow = 2
    For r = 2 To lastRow
        wsMaintain.Cells(outRow + 3, 2).NumberFormat = "@"
        wsMaintain.Cells(outRow + 3, 2).value = wsDest.Cells(r, 1).value
        outRow = outRow + 1
    Next r

Cleanup:
    On Error Resume Next
    wsDest.Visible = xlSheetHidden
    ActiveWorkbook.Protect Structure:=True, Windows:=True
    
    If Not wsStart Is Nothing Then wsStart.Activate
    
    Set wsStart = Nothing
    Set wsDest = Nothing
    Set wsMaintain = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "LoadManufacturerList error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
