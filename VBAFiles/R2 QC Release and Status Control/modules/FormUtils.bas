Attribute VB_Name = "FormUtils"
Option Explicit

'--- Unprotects the active worksheet (no password required)
Public Sub UnprotectActiveSheet()
    On Error Resume Next
    ActiveSheet.Unprotect
    On Error GoTo 0
End Sub

'--- Protects the active worksheet (no password required)
Public Sub ProtectActiveSheet()
    On Error Resume Next
    If PROD_DEV <> "DEV" Then
        ActiveSheet.Protect _
            AllowFiltering:=False, _
            AllowSorting:=False
    End If
    On Error GoTo 0
End Sub

Public Function GetFieldValByLabelC(ByVal labelText As String) As String
    ' Finds the label in column C (any row) and returns the value from column D, same row.
    ' Match is case-insensitive and ignores surrounding spaces.
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    Dim r As Long
    For r = 1 To lastRow
        If LCase$(Trim$(ws.Cells(r, "C").value)) = LCase$(Trim$(labelText)) Then
            GetFieldValByLabelC = Trim$(ws.Cells(r, "D").value)
            Exit Function
        End If
    Next r
    GetFieldValByLabelC = ""
End Function

Public Function YesNoToYN(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    If t = "yes" Then
        YesNoToYN = "Y"
    ElseIf t = "no" Then
        YesNoToYN = "N"
    Else
        ' If blank or unrecognized, default to "N" (adjust if you prefer blank)
        YesNoToYN = "N"
    End If
End Function

'===== Helper: map truthy values to Yes/No for the form =====
Public Function ToYesNo(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    Select Case t
        Case "1", "-1", "y", "yes", "true"
            ToYesNo = "Yes"
        Case "0", "n", "no", "false", ""
            ToYesNo = "No"
        Case Else
            ' If unrecognized, default to "No"
            ToYesNo = "Error"
    End Select
End Function

' --- Append QC Number record to audit trail file ---
Public Sub AppendToQCNumberAudit(ByVal recordValues As Variant)
    Dim auditPath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    
    On Error GoTo ErrorHandler
    
    ' Define audit file path
    auditPath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Format record for CSV
    If IsArray(recordValues) Then
        line = Join(recordValues, ",")
    Else
        line = recordValues ' In case a pre-formatted string is passed
    End If
    
    ' Append record to audit file
    Set ts = fso.OpenTextFile(auditPath, 8, True) ' 8 = ForAppending
    ts.WriteLine line
    ts.Close
    
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error writing to audit trail: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
End Sub

Public Function ValidateNoCommas(ByVal value As String, ByVal fieldName As String) As Boolean
    If InStr(1, value, ",") > 0 Then
        MsgBox "Error: " & fieldName & " cannot contain commas. Please remove any commas and try again.", vbCritical
        ValidateNoCommas = False
    Else
        ValidateNoCommas = True
    End If
End Function