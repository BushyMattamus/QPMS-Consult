Attribute VB_Name = "CSVUtils"
Option Explicit

' Imports a CSV file into the specified worksheet and range using QueryTable.
' Returns True if successful, False otherwise.
Public Function ImportCSVWithQueryTable(ByVal csvPath As String, ByVal ws As Worksheet, ByVal destRange As Range, Optional ByVal clearSheet As Boolean = True) As Boolean
    Dim qt As QueryTable
    On Error GoTo ErrorHandler

    If clearSheet Then
        ws.Cells.Clear
    End If

    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=destRange)
    With qt
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001                 ' UTF-8
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileTrailingMinusNumbers = True
        .PreserveFormatting = True
        .AdjustColumnWidth = False
        .RefreshStyle = xlOverwriteCells
        .TextFileColumnDataTypes = Array(xlTextFormat)
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    Set qt = Nothing
    Set destRange = Nothing
    Set ws = Nothing

    ImportCSVWithQueryTable = True
    Exit Function

ErrorHandler:
    ImportCSVWithQueryTable = False
End Function

Public Function ImportCSVWithOpenText( _
    ByVal csvPath As String, _
    ByVal ws As Worksheet, _
    ByVal destRange As Range, _
    Optional ByVal clearSheet As Boolean = True, _
    Optional ByVal sourceFolderPath As String = "" _
) As Boolean

    On Error GoTo ErrHandler
    
    Dim fso As Object
    Dim tempCsvPath As String
    Dim csvWb As Workbook
    Dim csvWs As Worksheet
    Dim dataRange As Range
    Dim colCount As Long
    Dim fieldInfo As Variant
    Dim i As Long
    Dim ts As Object
    Dim firstLine As String
    Dim commaCount As Long
    Dim sourceFileName As String
    
    ImportCSVWithOpenText = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    sourceFileName = fso.GetFileName(csvPath)
    
    If sourceFolderPath = "" Then
        tempCsvPath = CreateTemporaryFileCopy(sourceFileName)
    Else
        tempCsvPath = CreateTemporaryFileCopy(sourceFileName, sourceFolderPath)
    End If
    
    If tempCsvPath = "" Then
        MsgBox "Failed to create temporary copy of " & sourceFileName, vbCritical, "ImportCSVWithOpenText"
        GoTo CleanExit
    End If
    
    If clearSheet Then
        ws.Cells.Clear
    End If
    
    Set ts = fso.OpenTextFile(tempCsvPath, 1, False)
    If Not ts.AtEndOfStream Then
        firstLine = ts.ReadLine
    End If
    ts.Close
    Set ts = Nothing
    
    commaCount = Len(firstLine) - Len(Replace(firstLine, ",", ""))
    colCount = commaCount + 1
    
    ReDim fieldInfo(1 To colCount)
    For i = 1 To colCount
        fieldInfo(i) = Array(i, xlTextFormat)
    Next i
    
    Workbooks.OpenText _
        fileName:=tempCsvPath, _
        DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Comma:=True, _
        fieldInfo:=fieldInfo
    
    Set csvWb = ActiveWorkbook
    Set csvWs = csvWb.Worksheets(1)
    
    Set dataRange = csvWs.UsedRange
    destRange.Resize(dataRange.Rows.Count, dataRange.Columns.Count).value = dataRange.value
    
    csvWb.Close SaveChanges:=False
    
    ImportCSVWithOpenText = True
    GoTo CleanExit
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ImportCSVWithOpenText"
    On Error Resume Next
    If Not csvWb Is Nothing Then csvWb.Close SaveChanges:=False
    GoTo CleanExit
    
CleanExit:
    On Error Resume Next
    If Len(tempCsvPath) > 0 Then
        If fso.FileExists(tempCsvPath) Then fso.DeleteFile tempCsvPath
    End If
    Set ts = Nothing
    Set dataRange = Nothing
    Set csvWs = Nothing
    Set csvWb = Nothing
    Set destRange = Nothing
    Set ws = Nothing
    Set fso = Nothing
    Exit Function
End Function

Public Function ExportCSVToExcel(ByVal csvName As String, ByVal exportPrefix As String) As Boolean
    Const PROC_NAME As String = "ExportCSVToExcel"
    Dim srcPath As String
    Dim destPath As String
    Dim wb As Workbook
    
    On Error GoTo Cleanup
    
    srcPath = BuildFilePath(GetWorkbookPath(), FOLDER_TABLES, csvName)
    destPath = BuildExportPath(exportPrefix & GenerateTimestamp() & ".xlsx")
    
    EnsureFolderExists BuildFilePath(GetWorkbookPath(), FOLDER_EXPORTS)
    
    DisableAppUpdates
    Set wb = Workbooks.Open(fileName:=srcPath, ReadOnly:=True)
    
    wb.SaveAs fileName:=destPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    
    MsgBox "Export completed:" & vbCrLf & destPath, vbInformation
    ExportCSVToExcel = True
    GoTo Cleanup
    
Cleanup:
    RestoreAppState
    Exit Function
End Function

Public Function ReadCSVFile(ByVal filePath As String) As String()
    Const PROC_NAME As String = "ReadCSVFile"
    Dim fileContent As String
    Dim lines() As String
    
    On Error GoTo ErrorHandler
    
    DisableAppUpdates
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Type = 2 'adTypeText
        .Open
        .LoadFromFile filePath
        fileContent = .ReadText
        .Close
    End With
    
    ' Ensure we return a properly dimensioned array
    If Len(Trim$(fileContent)) = 0 Then
        ReDim lines(0 To 0)
        lines(0) = ""
    Else
        lines = Split(ToLF(fileContent), vbLf)
    End If
    
    ReadCSVFile = lines
    GoTo Cleanup
    
ErrorHandler:
    ReDim lines(0 To 0)
    lines(0) = ""
    ReadCSVFile = lines
    
Cleanup:
    RestoreAppState
    Exit Function
End Function

Public Function ParseCSVLine(ByVal line As String) As String()
    Dim elements() As String
    Dim inQuotes As Boolean
    Dim currentElement As String
    Dim i As Long
    Dim char As String
    Dim result() As String
    Dim resultCount As Long
    
    ReDim result(0)
    resultCount = 0
    
    For i = 1 To Len(line)
        char = Mid(line, i, 1)
        
        Select Case char
            Case """"
                inQuotes = Not inQuotes
            Case ","
                If Not inQuotes Then
                    ReDim Preserve result(resultCount)
                    result(resultCount) = currentElement
                    resultCount = resultCount + 1
                    currentElement = ""
                Else
                    currentElement = currentElement & char
                End If
            Case Else
                currentElement = currentElement & char
        End Select
    Next i
    
    ' Add last element
    ReDim Preserve result(resultCount)
    result(resultCount) = currentElement
    
    ParseCSVLine = result
End Function

Public Function ToLF(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    ToLF = s
End Function

' Helper function to write transactions to RM_Trans.csv with file locking
Public Sub WriteRMTransaction(transType As String, qcNumber As String, qcStat As String, quantity As Double, masterUoM As String, qcTransID As String, Optional ByVal adjNote As String = "")
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim transID As String
    
    On Error GoTo WriteRMError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = GetWorkbookTableFilePath(FILE_RM_TRANS)
    
    ' Get next transaction ID
    transID = GetNextTransactionID()
    
    ' Build CSV line with proper quoting - includes date AND time and QC_Trans_ID
    Dim line As String
    line = """" & transID & """," & _
           """" & transType & """," & _
           """" & qcNumber & """," & _
           """" & qcStat & """," & _
           """""," & _
           """""," & _
           """" & CStr(Round(quantity, 3)) & """," & _
           """" & masterUoM & """," & _
           """" & Environ("Username") & """," & _
           """" & GetTransactionTimestamp() & """," & _
           """" & qcTransID & """" & _
           """""," & _
           """" & adjNote & """," & _
           """""," & _
           """""," & _
           """"","
    
    ' Append to RM_Trans.csv
    Set ts = fso.OpenTextFile(filePath, 8, True)  ' 8 = ForAppending
    ts.WriteLine line
    ts.Close
    
    GoTo WriteRMCleanup

WriteRMError:
    MsgBox "Error writing to RM_Trans.csv: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close

WriteRMCleanup:
    Set ts = Nothing
    Set fso = Nothing
End Sub
