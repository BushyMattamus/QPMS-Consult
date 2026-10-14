Attribute VB_Name = "BusLogNavigation"
Sub GoTo_tblProdOutput()
    ' Navigates to the worksheet named "tblProdOutput"
    On Error Resume Next
    ThisWorkbook.Sheets("tblProdOutput").Activate
    On Error GoTo 0
End Sub

Sub GoTo_Products()
    ' Navigates to the worksheet named "tblCPMU_Product_List"
    On Error Resume Next
    ThisWorkbook.Sheets("tblCPMU_Product_List").Activate
    On Error GoTo 0
End Sub

Sub GoTo_Lists()
    ' Navigates to the worksheet named "tblProdOutput"
    On Error Resume Next
    ThisWorkbook.Sheets("Table Lists").Activate
    On Error GoTo 0
End Sub

Sub GoTo_Entry()
    ' Navigates to the worksheet named "tblProdOutput"
    On Error Resume Next
    ThisWorkbook.Sheets("Production Entry").Activate
    On Error GoTo 0
End Sub

Sub GoTo_Email()
    ' Navigates to the worksheet named "tblProdOutput"
    On Error Resume Next
    ThisWorkbook.Sheets("Email").Activate
    On Error GoTo 0
End Sub

