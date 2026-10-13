Attribute VB_Name = "test2"
Option Explicit

Public Sub CalculateDiscount()
    ' Sample logic for testing:
    ' Cell A2 = Quantity
    ' Cell B2 = Unit Price
    ' Cell C2 = Discounted Total
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim qty As Double
    Dim unitPrice As Double
    Dim total As Double
    
    qty = CDbl(ws.Range("A2").Value)
    unitPrice = CDbl(ws.Range("B2").Value)
    
    ' If qty > 20, apply 15% discount
    If qty > 20 Then
        total = qty * unitPrice * 0.85
    Else
        total = qty * unitPrice
    End If
    
    ws.Range("C2").Value = total
End Sub