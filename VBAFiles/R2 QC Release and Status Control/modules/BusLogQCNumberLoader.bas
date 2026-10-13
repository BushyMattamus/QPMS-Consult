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
    Dim lastRetestDate As String
    Dim firstRetestDate As String
    Dim retestDateValue As String
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

    Range("D4:D25").ClearContents

    Range("D4").value = Trim$(GetFieldValue(rs, HEADER_STANDARD_NAME) & "")
    Range("D5").value = Format$(GetFieldValue(rs, HEADER_RM_CODE) & "", "0000")
    Range("D6").value = Trim$(GetFieldValue(rs, HEADER_BRAND_NAME) & "")
    Range("D7").value = Trim$(GetFieldValue(rs, HEADER_QUICK_CODE) & "")
    Range("D8").value = Trim$(GetFieldValue(rs, HEADER_PART_ARTG) & "")
    Range("D9").value = Trim$(GetFieldValue(rs, HEADER_VENDOR_NAME) & "")
    Range("D10").value = Trim$(GetFieldValue(rs, HEADER_MANUFACTURER_NAME) & "")
    Range("D11").value = Trim$(GetFieldValue(rs, HEADER_MANUFACTURER_LOT) & "")
    SetDateStringInCell Range("D12"), ANSIToDate(Trim$(GetFieldValue(rs, HEADER_MANUFACTURER_EXPIRY) & ""))
    Range("D13").value = Trim$(GetFieldValue(rs, HEADER_NUMBER_CONTAINERS) & "")
    Range("D15").value = Trim$(GetFieldValue(rs, HEADER_MASTER_UOM) & "")
    Range("D16").value = Trim$(GetFieldValue(rs, HEADER_QC_STATUS_FLAG) & "")
    Range("D17").value = Trim$(GetFieldValue(rs, HEADER_QC_TESTED) & "")
    SetDateStringInCell Range("D18"), ANSIToDate(Trim$(GetFieldValue(rs, HEADER_TESTED_DATE) & ""))
    Range("D19").value = Trim$(GetFieldValue(rs, HEADER_NEW_RETEST) & "")

    lastRetestDate = ANSIToDate(Trim$(GetFieldValue(rs, HEADER_LAST_RETEST_DATE) & ""))
    firstRetestDate = ANSIToDate(Trim$(GetFieldValue(rs, HEADER_FIRST_RETEST_DATE) & ""))

    currentBalance = GetQCBalanceFromRMTrans(qcNumber)
    Range("D14").value = currentBalance

    If Len(Trim$(lastRetestDate)) > 0 Then
        retestDateValue = lastRetestDate
    Else
        retestDateValue = firstRetestDate
    End If
    
    If Len(Trim$(retestDateValue)) > 0 Then
        SetDateStringInCell Range("D20"), retestDateValue
    Else
        Range("D20").value = ""
    End If

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
        If FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Sub
    
End Sub

Public Sub LoadQCNumberForRelease()
    ' Get the QC number
    Dim qcNumber As String
    qcNumber = GetFieldValByLabelC("QC Number:")
    
    If qcNumber = "" Then
        MsgBox "Please enter a QC Number in cell D3", vbExclamation
        Exit Sub
    End If
    
    ' Get the file path
    Dim csvPath As String
    csvPath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not FileExists(csvPath) Then
        MsgBox "QC Number database file not found: " & csvPath, vbExclamation
        Exit Sub
    End If
    
    ' Read and process the CSV file
    Dim fileNum As Integer
    Dim headerLine As String
    Dim dataLine As String
    Dim headers() As String
    Dim values() As String
    Dim foundRecord As Boolean
    Dim i As Integer
    
    fileNum = FreeFile
    Open csvPath For Input As #fileNum
    
    ' Read headers
    Line Input #fileNum, headerLine
    headers = Split(headerLine, ",")
    
    ' Search for matching QC number
    foundRecord = False
    Do Until EOF(fileNum)
        Line Input #fileNum, dataLine
        values = Split(dataLine, ",")
        
        ' Check if this is our QC number (first column)
        If Len(values(0)) > 0 Then
            If Trim$(Replace(values(0), """", "")) = qcNumber Then
                foundRecord = True
                Exit Do
            End If
        End If
    Loop
    
    Close #fileNum
    
    If Not foundRecord Then
        MsgBox "QC Number " & qcNumber & " not found in database.", vbExclamation
        Exit Sub
    End If
    
    ' Check QC Status Flag before loading
    Dim qcStatusValue As String
    Dim qcStatusIndex As Integer
    
    ' Find QC Status Flag index
    For i = 0 To UBound(headers)
        If Trim$(headers(i)) = HEADER_QC_STATUS_FLAG Then
            qcStatusIndex = i
            Exit For
        End If
    Next i
    
    ' Get QC Status value
    If qcStatusIndex <= UBound(values) Then
        qcStatusValue = Trim$(Replace(values(qcStatusIndex), """", ""))
    End If
    
    ' Validate QC Status Flag
    If UCase$(qcStatusValue) <> "SAMPLED" And UCase$(qcStatusValue) <> "QUARANTINE" Then
        Dim warningMsg As String
        warningMsg = "Warning: QC Number " & qcNumber & " has QC Status Flag '" & qcStatusValue & "'." & vbCrLf & vbCrLf & _
                    "This record is not in 'Sampled' or 'Quarantine' status." & vbCrLf & vbCrLf & _
                    "Click OK to proceed with loading this record, or Cancel to abort."
        
        If MsgBox(warningMsg, vbOKCancel + vbExclamation, "QC Status Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    
    ' Populate the form fields
    
    ' Map the values to their respective fields
    Dim lastRetestDate As String
    Dim firstRetestDate As String
        
    For i = 0 To UBound(headers)
        If i > UBound(values) Then Exit For   ' safety guard
        
        Dim cleanValue As String
        cleanValue = Replace(values(i), """", "")
        
        Select Case Trim$(headers(i))
            Case HEADER_STANDARD_NAME
                Range("D4").value = cleanValue
            Case HEADER_RM_CODE
                Range("D5").value = Format$(cleanValue, "0000")
            Case HEADER_MANUFACTURER_NAME
                Range("D6").value = cleanValue
            Case HEADER_MANUFACTURER_LOT
                Range("D7").value = cleanValue
            Case HEADER_MANUFACTURER_EXPIRY
                ' Use helper function for date display
                SetDateStringInCell Range("D8"), ANSIToDate(cleanValue)
            Case HEADER_LAST_RETEST_DATE
                lastRetestDate = ANSIToDate(cleanValue)
            Case HEADER_FIRST_RETEST_DATE
                firstRetestDate = ANSIToDate(cleanValue)
            Case HEADER_QC_STATUS_FLAG
                Range("D9").value = cleanValue
        End Select
    Next i
    
    ' Handle retest date logic - use last retest if available, otherwise first retest
    Dim retestDateValue As String
    If Len(Trim$(lastRetestDate)) > 0 Then
        retestDateValue = lastRetestDate
    Else
        retestDateValue = firstRetestDate
    End If

    ' Set "New or Retest" field
    If Len(Trim$(lastRetestDate)) = 0 And Len(Trim$(firstRetestDate)) = 0 Then
        Range("D12").value = "New"
    Else
        Range("D12").value = "Retest"
    End If

    ' Set the retest date - already in correct format from ANSIToDate
    If Len(Trim$(retestDateValue)) > 0 Then
        SetDateStringInCell Range("D13"), retestDateValue
    Else
        Range("D13").value = ""
    End If

End Sub



