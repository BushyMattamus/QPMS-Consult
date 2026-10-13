Attribute VB_Name = "ProductLoad"
Option Explicit

Public Sub LoadProductFields(sheetName As String)

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim wsRef As Worksheet
    Dim descCell As Range
    Dim productDesc As String
    Dim rowNum As Variant
    Dim shelfNo As String
    Dim shelfLife As String
    Dim missing As String

    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsRef = ThisWorkbook.Sheets("MasterRef")

    ' Get the cell beside the Product Description label
    Set descCell = GetFieldByLabel(sheetName, "C", "Product Description")
    productDesc = descCell.value

   
    ' If empty, clear all associated fields
    If productDesc = "" Then
        ClearProductFields
        Exit Sub
    End If

    ' Look for Product Description in column 3 of MasterRef
    rowNum = Application.Match(productDesc, wsRef.Columns(3), 0)

    ' If not found ? clear fields
    If IsError(rowNum) Or rowNum = 0 Then
        ClearProductFields
        Exit Sub
    End If

    ' Populate fields
    GetFieldByLabel(sheetName, "C", "PMI number").value = wsRef.Cells(rowNum, 1).value
    GetFieldByLabel(sheetName, "C", "PMI version").value = wsRef.Cells(rowNum, 2).value
    GetFieldByLabel(sheetName, "C", "Shelf life(No of days/Months/Years)").value = wsRef.Cells(rowNum, 6).value
    GetFieldByLabel(sheetName, "C", "Shelf life(Days/Months/Years)").value = wsRef.Cells(rowNum, 5).value
    GetFieldByLabel("Batch Number Generation", "C", "Manufacturing date").value = Date ' Format(Date, "dd/mm/yyyy")
    
    ' Validate shelf life fields
    shelfNo = GetFieldByLabel(sheetName, "C", "Shelf life(No of days/Months/Years)").value
    shelfLife = GetFieldByLabel(sheetName, "C", "Shelf life(Days/Months/Years)").value
    
    missing = ""
    
    If UCase(Trim(shelfLife)) <> "AS PER MANUFACTURER" Then
        If Trim(shelfNo) = "" Then missing = missing & vbCrLf & "- Shelf Life (Number)"
    End If
    
    If Trim(shelfLife) = "" Then missing = missing & vbCrLf & "- Shelf Life (Type)"
    
    If missing <> "" Then
        MsgBox "Warning: The following shelf life information is missing for this product:" & missing, vbExclamation, "Missing Shelf Life Data"
    End If
    
    ' Recalculate expiry date
    Call CalculateExpiryDate("Batch Number Generation")
    Exit Sub

ErrHandler:
    MsgBox "Error in UpdateProductFields: " & Err.Description, vbCritical

End Sub