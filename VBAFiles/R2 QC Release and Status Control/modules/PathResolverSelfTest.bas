Attribute VB_Name = "PathResolverSelfTest"
Option Explicit

Public Sub RunPathResolverSelfTest()
    Dim ws As Worksheet
    Dim tempCopy As String

    Set ws = ThisWorkbook.Worksheets(1)
    ws.Activate
    ws.Unprotect

    ws.Range("XFD1").value = GetDataEnvironment()
    ws.Range("XFD2").value = PROD_DEV()
    ws.Range("XFD3").value = IIf(FileExists(GetTableFilePath(FILE_RM_TRANS)), "OK", "MISSING")
    ws.Range("XFD4").value = IIf(FileExists(GetTableFilePath(FILE_QC_NUMBER)), "OK", "MISSING")
    ws.Range("XFD5").value = IIf(InStr(1, GetTableFilePath(FILE_RM_TRANS), "Material Management", vbTextCompare) > 0, "OK", "WRONG")

    tempCopy = CreateTemporaryFileCopy(FILE_RM_TRANS)
    ws.Range("XFD6").value = IIf(Len(tempCopy) > 0 And FileExists(tempCopy), "OK", "FAILED")
    If Len(tempCopy) > 0 Then Kill tempCopy

    ws.Protect
End Sub