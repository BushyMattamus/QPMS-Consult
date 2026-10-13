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
            AllowFiltering:=True, _
            AllowSorting:=True
    End If
    On Error GoTo 0
End Sub

Public Function ValidateNoCommas(ByVal value As String, ByVal fieldName As String) As Boolean
    If InStr(1, value, ",") > 0 Then
        MsgBox "Error: " & fieldName & " cannot contain commas. Please remove any commas and try again.", vbCritical
        ValidateNoCommas = False
    Else
        ValidateNoCommas = True
    End If
End Function
