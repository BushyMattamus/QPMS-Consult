Attribute VB_Name = "Access_ImportRawMaterial"
'=============================================================================
' Module: Access_ImportRawMaterial
' Purpose: Import RawMaterialMaster and RMChanges CSV files into Access
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: ImportRawMaterialMaster
' Returns: Number of records imported
'------------------------------------------------------------------------------
Public Function ImportRawMaterialMaster(dbPath As String) As Long
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim fso As Object
    Dim csvFile As Object
    Dim lines() As String
    Dim fieldNames() As String
    Dim i As Integer, j As Integer
    Dim line As String
    Dim recordCount As Long
    Dim values() As String
    Dim sqlFields() As String
    Dim sqlValues() As String
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open GetDatabasePath()
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open GetDatabasePath()
    End If
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    Dim sql As String
    Dim insertErrors As Long, firstInsertError As String
    sql = "DELETE FROM [RawMaterialMaster]"
    conn.Execute sql
    
    ' Read CSV file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set csvFile = fso.OpenTextFile(GetCSVDataDirectory & "Raw_Material_Master.csv", 1, False, -2)
    
    ' Read header
    line = csvFile.ReadLine
    fieldNames = Split(Trim(line), ",")
    ReDim sqlFields(UBound(fieldNames))
    ReDim sqlValues(UBound(fieldNames))
    
    recordCount = 0
    
    ' Read data lines and build INSERT statements in memory
    Do While Not csvFile.AtEndOfStream
        line = csvFile.ReadLine
        
        If Trim(line) <> "" Then
            values = Split(Trim(line), ",")
            
            ' Ensure enough fields
            If UBound(values) < UBound(fieldNames) Then
                ReDim Preserve values(UBound(fieldNames))
            End If
            
            ' Build field list
            For j = 0 To UBound(fieldNames)
                sqlFields(j) = "[" & fieldNames(j) & "]"
            Next j
            
            ' Build value list
            For j = 0 To UBound(values)
                Dim fieldValue As String
                fieldValue = values(j)
                
                If IsNumeric(fieldValue) Or fieldValue = "" Then
                    If fieldValue = "" Then sqlValues(j) = "NULL" Else sqlValues(j) = fieldValue
                Else
                    fieldValue = Replace(fieldValue, "'", "''")
                    sqlValues(j) = "'" & fieldValue & "'"
                End If
            Next j
            
            ' Build complete INSERT statement using Join
            Dim fieldList As String
            Dim valueList As String
            fieldList = Join(sqlFields, ", ")
            valueList = Join(sqlValues, ", ")
            
            valueList = ApplyYesNoLiterals(valueList)
            sql = "INSERT INTO [RawMaterialMaster] (" & fieldList & ") VALUES (" & valueList & ")"
            
            On Error Resume Next
            conn.Execute sql
            If Err.Number = 0 Then
                recordCount = recordCount + 1
            Else
                insertErrors = insertErrors + 1
                If insertErrors = 1 Then firstInsertError = Err.Number & " " & Err.Description
            End If
            On Error GoTo ErrorHandler
        End If
    Loop
    
    csvFile.Close
    Set csvFile = Nothing
    conn.Close
    Set conn = Nothing
    
    If insertErrors > 0 Then PathResolver.LogTestMessage insertErrors & " insert(s) failed. First error: " & firstInsertError, "Import RawMaterialMaster"
    PathResolver.LogTestMessage "Imported " & recordCount & " row(s)", "Import RawMaterialMaster"
    ImportRawMaterialMaster = recordCount
    
    Exit Function
    
ErrorHandler:
    PathResolver.LogTestMessage "Failed " & Err.Number & " " & Err.Description, "Import RawMaterialMaster"
    ImportRawMaterialMaster = -1
End Function

'------------------------------------------------------------------------------
' Function: ImportRMChanges
' Returns: Number of records imported
'------------------------------------------------------------------------------
Public Function ImportRMChanges(dbPath As String) As Long
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim fso As Object
    Dim csvFile As Object
    Dim lines() As String
    Dim fieldNames() As String
    Dim i As Integer, j As Integer
    Dim line As String
    Dim recordCount As Long
    Dim values() As String
    Dim sqlFields() As String
    Dim sqlValues() As String
    
    ' Create connection
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open GetDatabasePath()
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open GetDatabasePath()
    End If
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    Dim sql As String
    Dim insertErrors As Long, firstInsertError As String
    sql = "DELETE FROM [RMChanges]"
    conn.Execute sql
    
    ' Read CSV file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set csvFile = fso.OpenTextFile(GetCSVDataDirectory & "RM_Master_Changes.csv", 1, False, -2)
    
    ' Read header
    line = csvFile.ReadLine
    fieldNames = Split(Trim(line), ",")
    ReDim sqlFields(UBound(fieldNames))
    ReDim sqlValues(UBound(fieldNames))
    
    recordCount = 0
    
    ' Read data lines and build INSERT statements in memory
    Do While Not csvFile.AtEndOfStream
        line = csvFile.ReadLine
        
        If Trim(line) <> "" Then
            values = Split(Trim(line), ",")
            
            ' Ensure enough fields
            If UBound(values) < UBound(fieldNames) Then
                ReDim Preserve values(UBound(fieldNames))
            End If
            
            ' Build field list
            For j = 0 To UBound(fieldNames)
                sqlFields(j) = "[" & fieldNames(j) & "]"
            Next j
            
            ' Build value list
            For j = 0 To UBound(values)
                Dim fieldValue As String
                fieldValue = values(j)
                
                If IsNumeric(fieldValue) Or fieldValue = "" Then
                    If fieldValue = "" Then sqlValues(j) = "NULL" Else sqlValues(j) = fieldValue
                Else
                    fieldValue = Replace(fieldValue, "'", "''")
                    sqlValues(j) = "'" & fieldValue & "'"
                End If
            Next j
            
            ' Build complete INSERT statement using Join
            Dim fieldList As String
            Dim valueList As String
            fieldList = Join(sqlFields, ", ")
            valueList = Join(sqlValues, ", ")
            
            valueList = ApplyYesNoLiterals(valueList)
            sql = "INSERT INTO [RMChanges] (" & fieldList & ") VALUES (" & valueList & ")"
            
            On Error Resume Next
            conn.Execute sql
            If Err.Number = 0 Then
                recordCount = recordCount + 1
            Else
                insertErrors = insertErrors + 1
                If insertErrors = 1 Then firstInsertError = Err.Number & " " & Err.Description
            End If
            On Error GoTo ErrorHandler
        End If
    Loop
    
    csvFile.Close
    Set csvFile = Nothing
    conn.Close
    Set conn = Nothing
    
    If insertErrors > 0 Then PathResolver.LogTestMessage insertErrors & " insert(s) failed. First error: " & firstInsertError, "Import RMChanges"
    PathResolver.LogTestMessage "Imported " & recordCount & " row(s)", "Import RMChanges"
    ImportRMChanges = recordCount
    
    Exit Function
    
ErrorHandler:
    PathResolver.LogTestMessage "Failed " & Err.Number & " " & Err.Description, "Import RMChanges"
    ImportRMChanges = -1
End Function

' Splits on ", " because CSV values contain no commas
Public Function ApplyYesNoLiterals(ByVal valueList As String) As String
    Dim parts() As String
    Dim k As Long
    parts = Split(valueList, ", ")
    For k = 0 To UBound(parts)
        Select Case UCase(parts(k))
            Case "'Y'": parts(k) = "True"
            Case "'N'": parts(k) = "False"
        End Select
    Next k
    ApplyYesNoLiterals = Join(parts, ", ")
End Function
