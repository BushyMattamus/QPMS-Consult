Attribute VB_Name = "BusLogQCNumberLoader"
Option Explicit

Public Sub LoadQCNumberRecord()
    On Error GoTo ErrorHandler

    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    Dim qcNumber As String
    Dim currentBalance As Double

    qcNumber = GetFieldValByLabelC("QC Number:")
    
    If qcNumber = "" Then
        MsgBox "Please enter a QC Number in cell D3", vbExclamation
        GoTo CleanExit
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPath = CreateTemporaryFileCopy(FILE_QC_NUMBER)
    
    If tempPath = "" Then
        MsgBox "Failed to create temporary copy of QC Number file.", vbCritical, "LoadQCNumberRecord"
        GoTo CleanExit
    End If
    
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    Set conn = OpenADOConnection(folderPath)

    If conn Is Nothing Then
        MsgBox "Failed to open connection to temporary file.", vbCritical, "LoadQCNumberRecord"
        GoTo CleanExit
    End If

    sql = "SELECT * FROM [" & fileName & "]"

    Set rs = conn.Execute(sql)

    If rs.EOF Then
        MsgBox "QC Number database is empty.", vbExclamation
        GoTo CleanExit
    End If
    
    Dim foundRecord As Boolean
    foundRecord = False
    
    Do While Not rs.EOF
        If Trim$(GetFieldValue(rs, HEADER_QC_NUMBER) & "") = qcNumber Then
            foundRecord = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    
    If Not foundRecord Then
        MsgBox "QC Number " & qcNumber & " not found in database.", vbExclamation
        GoTo CleanExit
    End If

    Range("D4:D20").ClearContents

    Range("D4").value = Trim$(GetFieldValue(rs, HEADER_STANDARD_NAME) & "")
    Range("D5").value = Format$(GetFieldValue(rs, HEADER_RM_CODE) & "", "0000")
    Range("D6").value = Trim$(GetFieldValue(rs, HEADER_BRAND_NAME) & "")
    Range("D7").value = Trim$(GetFieldValue(rs, HEADER_QUICK_CODE) & "")
    Range("D8").value = Trim$(GetFieldValue(rs, HEADER_PART_ARTG) & "")
    Range("D9").value = Trim$(GetFieldValue(rs, HEADER_VENDOR_NAME) & "")
    Range("D10").value = Trim$(GetFieldValue(rs, HEADER_PO_NUMBER) & "")
    Range("D11").value = Trim$(GetFieldValue(rs, HEADER_MANUFACTURER_NAME) & "")
    Range("D12").value = Trim$(GetFieldValue(rs, HEADER_MANUFACTURER_LOT) & "")
    SetDateStringInCell Range("D13"), ANSIToDate(Trim$(GetFieldValue(rs, HEADER_MANUFACTURER_EXPIRY) & ""))
    Range("D14").value = Trim$(GetFieldValue(rs, HEADER_NUMBER_CONTAINERS) & "")
    Range("D15").value = Trim$(GetFieldValue(rs, HEADER_TOTAL_QTY) & "")
    Range("D17").value = Trim$(GetFieldValue(rs, HEADER_MASTER_UOM) & "")
    Range("D18").value = Trim$(GetFieldValue(rs, HEADER_ITEM_COST) & "")
    Range("D19").value = Trim$(GetFieldValue(rs, HEADER_QC_STATUS_FLAG) & "")

    currentBalance = GetQCBalanceFromRMTrans(qcNumber)
    Range("D16").value = currentBalance

    GoTo CleanExit

ErrorHandler:
    MsgBox "LoadQCNumberRecord error: " & Err.Description, vbCritical
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not fso Is Nothing Then
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Sub
    
End Sub
