Attribute VB_Name = "Access_ImportMasterData"
Option Explicit

Public Function ImportMasterDataFinishedGoods(dbPath As String) As Long
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection, fso As Object, csvFile As Object
    Dim line As String, recordCount As Long
    Dim fieldNames() As String, values() As Variant
    Dim j As Integer, fieldValue As String
    Dim sql As String
    Dim fieldsStr As String, valsStr As String
    
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open GetDatabasePath()
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open GetDatabasePath()
    End If
    On Error GoTo ErrorHandler
    
    sql = "DELETE FROM [MasterDataFinishedGoods]"
    conn.Execute sql
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set csvFile = fso.OpenTextFile(GetCSVDataDirectory & "MasterDataFinishedGoods.csv", 1, False, -2)
    
    line = csvFile.ReadLine
    fieldNames = Split(Trim(line), ",")
    recordCount = 0
    
    Do While Not csvFile.AtEndOfStream
        line = csvFile.ReadLine
        
        If Trim(line) <> "" Then
            values = Split(Trim(line), ",")
            
            fieldsStr = ""
            For j = 0 To UBound(fieldNames)
                If j > 0 Then fieldsStr = fieldsStr & ", "
                fieldsStr = fieldsStr & "[" & fieldNames(j) & "]"
            Next j
            
            valsStr = ""
            For j = 0 To UBound(values)
                If j > 0 Then valsStr = valsStr & ", "
                
                fieldValue = values(j)
                
                If IsNumeric(fieldValue) Or fieldValue = "" Then
                    If fieldValue = "" Then valsStr = valsStr & "NULL" Else valsStr = valsStr & fieldValue
                Else
                    fieldValue = Replace(fieldValue, "'", "''")
                    valsStr = valsStr & "'" & fieldValue & "'"
                End If
            Next j
            
            sql = "INSERT INTO [MasterDataFinishedGoods] (" & fieldsStr & ") VALUES (" & valsStr & ")"
            
            On Error Resume Next
            conn.Execute sql
            If Err.Number = 0 Then recordCount = recordCount + 1
            On Error GoTo ErrorHandler
        End If
    Loop
    
    csvFile.Close
    Set csvFile = Nothing
    conn.Close
    Set conn = Nothing
    
    ImportMasterDataFinishedGoods = recordCount
    
    Exit Function
    
ErrorHandler:
    ImportMasterDataFinishedGoods = -1
End Function

Public Function ImportAuditTrail(dbPath As String) As Long
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection, fso As Object, csvFile As Object
    Dim line As String, recordCount As Long
    Dim fieldNames() As String, values() As Variant
    Dim j As Integer, fieldValue As String
    Dim sql As String
    Dim fieldsStr As String, valsStr As String
    
    Set conn = New ADODB.Connection
    conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    
    On Error Resume Next
    conn.Open GetDatabasePath()
    If Err.Number <> 0 Then
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        conn.Open GetDatabasePath()
    End If
    On Error GoTo ErrorHandler
    
    sql = "DELETE FROM [AuditTrail]"
    conn.Execute sql
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set csvFile = fso.OpenTextFile(GetCSVDataDirectory & "AuditTrailMasterDataFinishedGoods.csv", 1, False, -2)
    
    line = csvFile.ReadLine
    fieldNames = Split(Trim(line), ",")
    recordCount = 0
    
    Do While Not csvFile.AtEndOfStream
        line = csvFile.ReadLine
        
        If Trim(line) <> "" Then
            values = Split(Trim(line), ",")
            
            fieldsStr = ""
            For j = 0 To UBound(fieldNames)
                If j > 0 Then fieldsStr = fieldsStr & ", "
                fieldsStr = fieldsStr & "[" & fieldNames(j) & "]"
            Next j
            
            valsStr = ""
            For j = 0 To UBound(values)
                If j > 0 Then valsStr = valsStr & ", "
                
                fieldValue = values(j)
                
                If IsNumeric(fieldValue) Or fieldValue = "" Then
                    If fieldValue = "" Then valsStr = valsStr & "NULL" Else valsStr = valsStr & fieldValue
                Else
                    fieldValue = Replace(fieldValue, "'", "''")
                    valsStr = valsStr & "'" & fieldValue & "'"
                End If
            Next j
            
            sql = "INSERT INTO [AuditTrail] (" & fieldsStr & ") VALUES (" & valsStr & ")"
            
            On Error Resume Next
            conn.Execute sql
            If Err.Number = 0 Then recordCount = recordCount + 1
            On Error GoTo ErrorHandler
        End If
    Loop
    
    csvFile.Close
    Set csvFile = Nothing
    conn.Close
    Set conn = Nothing
    
    ImportAuditTrail = recordCount
    
    Exit Function
    
ErrorHandler:
    ImportAuditTrail = -1
End Function
