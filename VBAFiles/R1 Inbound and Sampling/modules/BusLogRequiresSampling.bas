Attribute VB_Name = "BusLogRequiresSampling"
Public Sub RequiresSampling()
    
    Dim wsSource As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim reportRange As Range
    Dim i As Long, outRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Refresh the QC numbers
    LoadQCNumberData
    
    ' Get source worksheet
    Set wsSource = ThisWorkbook.Worksheets("QC_Number")
    
    ' Get last row of data in source
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No QC Number data found. Please refresh QC Number data first.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Check if Summary sheet exists, create if not
    On Error Resume Next
    Set wsReport = ThisWorkbook.Worksheets("Requires_Sampling")
    On Error GoTo ErrorHandler
    
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReport.Name = "Requires_Sampling"
    End If
    
    UnprotectActiveSheet
    
    ' Clear existing content
    wsReport.Range("B:F").Clear
    
    ' Write headers
    wsReport.Range("B1:F1").value = Array( _
        HEADER_GR_DATE, _
        HEADER_QC_NUMBER, _
        HEADER_RM_CODE, _
        HEADER_STANDARD_NAME, _
        HEADER_QC_STATUS_FLAG)
    
    ' Copy filtered data (only "Receipted" status)
    outRow = 2
    For i = 2 To lastRow
        If LCase$(Trim$(wsSource.Cells(i, 34).value)) = "receipted" Then
            wsReport.Cells(outRow, 2).value = wsSource.Cells(i, 15).value  ' GR_Date
            wsReport.Cells(outRow, 3).value = wsSource.Cells(i, 1).value  ' QC_Number
            wsReport.Cells(outRow, 4).value = wsSource.Cells(i, 2).value  ' RM_Code
            wsReport.Cells(outRow, 5).value = wsSource.Cells(i, 3).value  ' Standard_Name
            wsReport.Cells(outRow, 6).value = wsSource.Cells(i, 34).value  ' QC_Status_Flag
            outRow = outRow + 1
        End If
    Next i
    
    ' Check if any records were found
    If outRow = 2 Then
        MsgBox "No records found with status 'Receipted'.", vbInformation
        GoTo CleanExit
    End If
    
    ' Format the report range
    Set reportRange = wsReport.Range("B1").Resize(outRow - 1, 5)
    With reportRange
        .Columns(1).NumberFormat = "dd/mm/yyyy"  ' Date format
        .Columns(2).NumberFormat = "@"           ' Text format for QC Number
        .Columns(3).NumberFormat = "0000"           ' Text format for RM Code
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Rows(1).Font.Bold = True
        .Locked = False
        .AutoFilter
        .EntireColumn.AutoFit
    End With
    
    wsReport.Range("B1:F1").Locked = False

    ' Activate the report sheet
    wsReport.Activate
    wsReport.Range("B1").Select
    
    MsgBox "Required Sampling report created successfully. Found " & (outRow - 2) & " records requiring sampling.", vbInformation
    
    ProtectActiveSheet
    
    GoTo CleanExit
    
ErrorHandler:
    MsgBox "Error creating Required Sampling: " & Err.Description, vbCritical
    
CleanExit:
End Sub

