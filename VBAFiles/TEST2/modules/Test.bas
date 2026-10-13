Attribute VB_Name = "Test"
Option Explicit

Public Sub CalculateDiscount()
    ' Sample logic for testing:
    ' Cell A1 = Quantity
    ' Cell B1 = Unit Price
    ' Cell C1 = Discounted Total
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim qty As Double
    Dim unitPrice As Double
    Dim total As Double
    
    qty = CDbl(ws.Range("A1").Value)
    unitPrice = CDbl(ws.Range("B1").Value)
    
    ' If qty > 10, apply 10% discount
    If qty > 10 Then
        total = qty * unitPrice * 0.9
    Else
        total = qty * unitPrice
    End If
    
    ws.Range("C1").Value = total
End Sub