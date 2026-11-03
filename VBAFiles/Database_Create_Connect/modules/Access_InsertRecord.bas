Attribute VB_Name = "Access_InsertRecord"
'=============================================================================
' Module: Access_InsertRecord
' Purpose: Insert single records into Access tables
' Project: Database_Create_Connect
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: InsertRecordIntoTable
' Purpose: Insert a single record into an Access table
' Returns: True if successful, False otherwise
'------------------------------------------------------------------------------
Public Function InsertRecordIntoTable(conn As ADODB.Connection, tableName As String, fieldNames() As String, values() As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim j As Integer
    Dim fieldValue As String
    Dim valueStr As String
    
    ' Build INSERT statement
    sql = "INSERT INTO [" & tableName & "] ("
    
    For j = 0 To UBound(fieldNames)
        If j > 0 Then sql = sql ", "
        sql = sql "[" & fieldNames(j) & "]"
    Next j
    
    sql = sql ") VALUES ("
    
    For j = 0 To UBound(values)
        If j > 0 Then sql = sql ", "
        
        fieldValue = values(j)
        
        ' Handle different data types
        If IsNumeric(fieldValue) Or fieldValue = "" Then
            ' Numeric or empty value
            If fieldValue = "" Then
                sql = sql "NULL"
            Else
                sql = sql fieldValue
            End If
        Else
            ' Text value - enclose in quotes and escape single quotes
            fieldValue = Replace(fieldValue, "'", "''")
            sql = sql "'" & fieldValue & "'"
        End If
    Next j
    
    sql = sql ")"
    
    ' Execute INSERT
    On Error Resume Next
    conn.Execute sql
    If Err.Number <> 0 Then
        InsertRecordIntoTable = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    InsertRecordIntoTable = True
    
    Exit Function
    
ErrorHandler:
    InsertRecordIntoTable = False
End Function
