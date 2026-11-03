Attribute VB_Name = "CSVReader"
'=============================================================================
' Module: CSVReader
' Purpose: Read CSV files into memory for processing
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: ReadCSVFile
' Returns: Object containing CSV data (lines array)
'------------------------------------------------------------------------------
Public Function ReadCSVFile(csvPath As String) As Object
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim csvFile As Object
    Dim lines() As String
    Dim i As Integer
    Dim line As String
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open CSV file
    Set csvFile = fso.OpenTextFile(csvPath, 1, False, -2) ' ForReading, Type: DelimitedText, DetectEOC
    
    ' Read all lines into array
    i = 0
    Do While Not csvFile.AtEndOfStream
        line = csvFile.ReadLine
        If line <> "" Then
            ReDim Preserve lines(i)
            lines(i) = line
            i = i + 1
        End If
    Loop
    
    ' Return the lines array
    ReadCSVFile = lines
    
    Exit Function
    
ErrorHandler:
    Set ReadCSVFile = Nothing
End Function

'------------------------------------------------------------------------------
' Function: GetCSVHeader
' Returns: Array of field names from CSV header line
'------------------------------------------------------------------------------
Public Function GetCSVHeader(csvPath As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim csvFile As Object
    Dim line As String
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open CSV file
    Set csvFile = fso.OpenTextFile(csvPath, 1, False, -2)
    
    ' Read header line
    line = csvFile.ReadLine
    
    ' Split into fields
    GetCSVHeader = Split(Trim(line), ",")
    
    csvFile.Close
    Set csvFile = Nothing
    
    Exit Function
    
ErrorHandler:
    GetCSVHeader = Array()
End Function

'------------------------------------------------------------------------------
' Function: GetCSVFieldCount
' Returns: Number of fields in the CSV file
'------------------------------------------------------------------------------
Public Function GetCSVFieldCount(csvPath As String) As Long
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim csvFile As Object
    Dim line As String
    Dim fieldCount As Long
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open CSV file
    Set csvFile = fso.OpenTextFile(csvPath, 1, False, -2)
    
    ' Read first non-empty line
    Do While Not csvFile.AtEndOfStream
        line = csvFile.ReadLine
        If Trim(line) <> "" Then
            fieldCount = UBound(Split(Trim(line), ",")) + 1
            Exit Do
        End If
    Loop
    
    csvFile.Close
    Set csvFile = Nothing
    
    GetCSVFieldCount = fieldCount
    
    Exit Function
    
ErrorHandler:
    GetCSVFieldCount = 0
End Function
