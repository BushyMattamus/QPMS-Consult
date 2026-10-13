Attribute VB_Name = "BusLogExport"
Option Explicit

Private Function ExportFilteredQC(ByVal csvName As String, ByVal exportPrefix As String) As Boolean
    ExportFilteredQC = ExportFilteredTable(csvName, exportPrefix, BuildQCFilterMap)
End Function

Private Function ExportFilteredTable(ByVal csvName As String, ByVal exportPrefix As String, ByVal filterMap As Variant) As Boolean
    Const PROC_NAME As String = "ExportFilteredTable"
    
    Dim tempPath As String
    Dim destPath As String
    Dim fileContent() As String
    Dim headers() As String
    Dim values() As String
    Dim headerMapping As Object
    Dim outputRows() As String
    Dim matchCount As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filters As Collection
    Dim i As Long, r As Long
    Dim exportFolder As String
    Dim fso As Object
    Dim balanceLookup As Object
    Dim sampleLookup As Object
    Dim qcBalanceIdx As Long
    Dim totalSampleQtyIdx As Long
    Dim qcNumberIdx As Long
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    tempPath = CreateTemporaryFileCopy(csvName)
    
    If tempPath = "" Then
        MsgBox "Failed to create temporary copy of " & csvName, vbCritical, PROC_NAME
        GoTo Cleanup
    End If
    
    exportFolder = BuildFilePath(GetWorkbookPath(), FOLDER_EXPORTS)
    
    If Not EnsureFolderExists(exportFolder) Then
    End If
    
    destPath = BuildExportPath(Environ("Username") & "_" & exportPrefix & EXCEL_FILE_EXT)
    
    fileContent = ReadCSVFile(tempPath)
    If UBound(fileContent) < 1 Then
        GoTo Cleanup
    End If
    
    headers = ParseCSVLine(fileContent(0))
    Set headerMapping = MapHeaders(headers)
    
    ' *** LOAD BALANCES AND SAMPLES FROM RM_TRANS ***
    Set balanceLookup = CreateObject("Scripting.Dictionary")
    Set sampleLookup = CreateObject("Scripting.Dictionary")
    Call LoadQCBalancesAndSamplesFromRMTrans(balanceLookup, sampleLookup)
    
    ' Find column indices for QC_Balance and Total_Sample_Qty
    qcBalanceIdx = -1
    qcNumberIdx = -1
    totalSampleQtyIdx = -1
    
    For i = 0 To UBound(headers)
        Dim headerClean As String
        headerClean = Trim$(Replace(headers(i), """", ""))
        If headerClean = HEADER_QC_BALANCE Then qcBalanceIdx = i
        If headerClean = HEADER_QC_NUMBER Then qcNumberIdx = i
        If headerClean = HEADER_TOTAL_SAMPLE_QTY Then totalSampleQtyIdx = i
    Next i
    
    ReDim outputRows(0 To UBound(fileContent))
    outputRows(0) = fileContent(0)
    
    Set filters = BuildFiltersFromMap(filterMap)
    
    matchCount = 0
    For i = 1 To UBound(fileContent)
        If Len(Trim$(fileContent(i))) > 0 Then
            values = ParseCSVLine(fileContent(i))
            If MatchesAllFilters(values, filters, headerMapping) Then
                
                ' *** UPDATE QC_BALANCE AND TOTAL_SAMPLE_QTY FROM RM_TRANS ***
                If qcNumberIdx >= 0 And qcNumberIdx <= UBound(values) Then
                    Dim currentQCNumber As String
                    currentQCNumber = Trim$(Replace(values(qcNumberIdx), """", ""))
                    
                    ' Update QC_Balance if column exists
                    If qcBalanceIdx >= 0 And qcBalanceIdx <= UBound(values) Then
                        If balanceLookup.Exists(currentQCNumber) Then
                            values(qcBalanceIdx) = """" & Format(balanceLookup(currentQCNumber), "0.000") & """"
                        Else
                            values(qcBalanceIdx) = """0.000"""
                        End If
                    End If
                    
                    ' Update Total_Sample_Qty if column exists
                    If totalSampleQtyIdx >= 0 And totalSampleQtyIdx <= UBound(values) Then
                        If sampleLookup.Exists(currentQCNumber) Then
                            values(totalSampleQtyIdx) = """" & Format(Abs(sampleLookup(currentQCNumber)), "0.000") & """"
                        Else
                            values(totalSampleQtyIdx) = """0.000"""
                        End If
                    End If
                    
                    ' Rebuild the CSV line with updated values
                    outputRows(matchCount + 1) = Join(values, ",")
                Else
                    outputRows(matchCount + 1) = fileContent(i)
                End If
                
                matchCount = matchCount + 1
            End If
        End If
    Next i
    
    If matchCount = 0 Then
        MsgBox "No matching records found.", vbInformation
        GoTo Cleanup
    End If
    
    ReDim Preserve outputRows(0 To matchCount)
    
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    
    For r = 0 To UBound(outputRows)
        values = ParseCSVLine(outputRows(r))
        ws.Range("A" & r + 1).Resize(1, UBound(values) + 1).value = values
    Next r
    
    ' Format the new worksheet
    With ws.Range("A1").Resize(matchCount + 1, UBound(ParseCSVLine(outputRows(0))) + 1)
        .Columns(2).NumberFormat = "@"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Rows(1).Font.Bold = True
        .AutoFilter
        .EntireColumn.AutoFit
    End With
        
    MsgBox "Export completed:" & vbCrLf & destPath & vbCrLf & _
           "Records found: " & matchCount, vbInformation
           
    ExportFilteredTable = True
    GoTo Cleanup
    
SaveAsError:
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
    MsgBox "SaveAs failed: " & Err.Description, vbCritical
    ExportFilteredTable = False
    Resume Cleanup

ErrorHandler:
    ExportFilteredTable = False
    If Err.Number <> 0 Then
        MsgBox "Error in " & PROC_NAME & ": " & Err.Description, vbCritical
    End If
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
    
Cleanup:
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set headerMapping = Nothing
    Set filters = Nothing
    Set balanceLookup = Nothing
    Set sampleLookup = Nothing
End Function

Private Function BuildFiltersFromMap(ByVal filterMap As Variant) As Collection
    Dim filters As New Collection
    Dim item As Variant
    Dim labelText As String
    Dim headerName As String
    Dim fieldValue As String
    
    For Each item In filterMap
        labelText = item(0)
        headerName = item(1)
        fieldValue = Trim$(GetFieldValByLabelC(labelText))
        AddFilterIfNotEmpty filters, headerName, fieldValue
    Next item
    
    Set BuildFiltersFromMap = filters
End Function

Private Function BuildQCFilterMap() As Variant
    BuildQCFilterMap = Array( _
        Array("QC Number:", HEADER_QC_NUMBER), _
        Array("Standard Name:", HEADER_STANDARD_NAME), _
        Array("RM Code:", HEADER_RM_CODE), _
        Array("Brand Name:", HEADER_BRAND_NAME), _
        Array("Quick Code:", HEADER_QUICK_CODE), _
        Array("Part Number/ARTG:", HEADER_PART_ARTG), _
        Array("Vendor Name:", HEADER_VENDOR_NAME), _
        Array("Manufactuer Name:", HEADER_MANUFACTURER_NAME), _
        Array("Manufactuer Lot Number:", HEADER_MANUFACTURER_LOT), _
        Array("VenManufactuerdor Expiry:", HEADER_MANUFACTURER_EXPIRY), _
        Array("No. Containers:", HEADER_NUMBER_CONTAINERS), _
        Array("Total Quantity:", HEADER_TOTAL_QTY), _
        Array("Unit of Measure:", HEADER_MASTER_UOM), _
        Array("QC Sample QTY:", HEADER_QC_SAMPLE_QTY), _
        Array("Retention Sample QTY:", HEADER_RETENTION_SAMPLE_QTY), _
        Array("Sample UoM:", HEADER_MASTER_UOM), _
        Array("QC Status:", HEADER_QC_STATUS_FLAG), _
        Array("Nil Stock:", HEADER_NIL_STOCK_FLAG) _
    )
End Function

Private Function MatchesAllFilters(ByRef values() As String, ByRef filters As Collection, ByRef headerMapping As Object) As Boolean
    Dim filter As Variant
    MatchesAllFilters = True
    
    For Each filter In filters
        Dim headerIndex As Long
        If headerMapping.Exists(filter(0)) Then
            headerIndex = headerMapping(filter(0))
            If headerIndex <= UBound(values) Then
                Dim fieldValue As String
                Dim filterValue As String
                fieldValue = Trim$(values(headerIndex))
                filterValue = Trim$(filter(1))
                
                Select Case filter(0)
                    Case HEADER_QC_NUMBER, HEADER_RM_CODE, HEADER_QC_STATUS_FLAG
                        If StrComp(fieldValue, filterValue, vbTextCompare) <> 0 Then
                            MatchesAllFilters = False
                            Exit Function
                        End If
                    Case Else
                        If InStr(1, fieldValue, filterValue, vbTextCompare) = 0 Then
                            MatchesAllFilters = False
                            Exit Function
                        End If
                End Select
            End If
        End If
    Next filter
End Function

Private Sub AddFilterIfNotEmpty(ByRef filters As Collection, ByVal headerName As String, ByVal filterValue As String)
    If Len(Trim$(filterValue)) > 0 Then
        filters.Add Array(headerName, filterValue)
    End If
End Sub

Public Sub ExportQCNumbers()
    If Not ExportFilteredQC(FILE_QC_NUMBER, PREFIX_HEADERS) Then
        MsgBox "Export failed", vbCritical
    End If
End Sub

Public Sub ExportQCNumberChanges()
    If Not ExportFilteredQC(FILE_QC_NUMBER_CHANGES, PREFIX_CHANGES) Then
        MsgBox "Export failed", vbCritical
    End If
End Sub

Public Sub ExportLiveReport()
    Const PROC_NAME As String = "ExportLiveReport"
    
    Dim wsReport As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim exportRange As Range
    Dim destPath As String
    Dim exportFolder As String
    
    On Error GoTo ErrorHandler
    
    Set wsReport = ActiveSheet
    
    ' Find last row with data in column F (QC Number column)
    lastRow = wsReport.Cells(wsReport.Rows.Count, 6).End(xlUp).row
    
    ' Check if there's any data to export (must be at least row 4 - header is row 3)
    If lastRow < 4 Then
        MsgBox "No data to export. Please generate the Live Report first.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Set the export range (F3 to N + last row) - 9 columns for live report
    Set exportRange = wsReport.Range("F3:N" & lastRow)
    
    ' Build export path
    exportFolder = BuildFilePath(GetWorkbookPath(), FOLDER_EXPORTS)
    
    ' Ensure export folder exists
    If Not EnsureFolderExists(exportFolder) Then
        ' Proceed but SaveAs may fail later
    End If
    
    destPath = BuildExportPath(Environ("Username") & "_Inbound_Live_Report" & EXCEL_FILE_EXT)
    
    ' Create new workbook
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "Inbound Live Report"
    
    ' Copy the data
    exportRange.Copy Destination:=ws.Range("A1")
    
    ' Format the new worksheet
    With ws.Range("A1").Resize(lastRow - 2, 9)
        .Columns(2).NumberFormat = "@"           ' RM Code as text
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Rows(1).Font.Bold = True
        .AutoFilter
        .EntireColumn.AutoFit
    End With
    
    ' Format RM Code column (B) with 0000 format
    With ws.Range("B2:B" & (lastRow - 2))
        .NumberFormat = "@"
        Dim i As Long
        For i = 2 To lastRow - 2
            If IsNumeric(ws.Cells(i, 2).value) Then
                ws.Cells(i, 2).value = Format(ws.Cells(i, 2).value, "0000")
            End If
        Next i
    End With
    
    ' Attempt SaveAs with guarded error handling
    On Error GoTo SaveAsError
    wb.SaveAs fileName:=destPath, FileFormat:=xlOpenXMLWorkbook
    On Error GoTo ErrorHandler
    
    wb.Close SaveChanges:=False
    
    MsgBox "Inbound Live Report exported successfully:" & vbCrLf & destPath & vbCrLf & _
           "Records exported: " & (lastRow - 3), vbInformation
           
    GoTo CleanExit
    
SaveAsError:
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
    MsgBox "SaveAs failed: " & Err.Description, vbCritical
    Resume CleanExit

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Error in " & PROC_NAME & ": " & Err.Description, vbCritical
    End If
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
    
CleanExit:
    Set ws = Nothing
    Set wb = Nothing
    Set wsReport = Nothing
    Set exportRange = Nothing
End Sub
