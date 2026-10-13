Attribute VB_Name = "BusLogMaintainQCStatus"
Option Explicit

Public Sub SaveQCRelease()
    
    ' Check for locked files first
    Dim qcFilePath As String, auditFilePath As String
    qcFilePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    auditFilePath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    
    If IsFileLocked(qcFilePath) Then
        MsgBox "Cannot proceed - QC Number file is in use by another user." & vbCrLf & _
               "Please try again when the file is available.", vbExclamation
        GoTo CleanExit
    End If
    
    If IsFileLocked(auditFilePath) Then
        MsgBox "Cannot proceed - QC Number Changes file is in use by another user." & vbCrLf & _
               "Please try again when the file is available.", vbExclamation
        GoTo CleanExit
    End If
    
    Dim mainWorkbook As Workbook
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim fileContent As String
    Dim lines() As String
    Dim line As String
    Dim lineParts() As String
    Dim found As Boolean
    Dim i As Long
    Dim qcTransID As String
    
    ' Declare required fields
    Dim qcNumber As String
    Dim StandardName As String
    Dim rmCode As String
    Dim mfgName As String
    Dim MfgLot As String
    Dim MfgExpiry As String
    Dim retestDate As String
    Dim QCStatus As String
    Dim NewRetest As String
    Dim qcTested As String
    Dim dateTested As String
    Dim previousQCStatus As String

    ' Get field values
    qcNumber = GetFieldValByLabelC("QC Number:")
    StandardName = GetFieldValByLabelC("Standard Name:")
    rmCode = Right$("0000" & Trim$(GetFieldValByLabelC("RM Code:")), 4)
    mfgName = GetFieldValByLabelC("Manufacturer Name:")
    MfgLot = GetFieldValByLabelC("Manufacturer Lot Number:")
    MfgExpiry = GetFieldValByLabelC("Manufacturer Expiry:")
    retestDate = GetFieldValByLabelC("Retest Date:")
    QCStatus = GetFieldValByLabelC("QC Status:")
    qcTested = GetFieldValByLabelC("QC Tested:")
    dateTested = GetFieldValByLabelC("Date Tested:")
    NewRetest = GetFieldValByLabelC("New or Retest:")

    If qcNumber = "" Then
        MsgBox "Please enter a QC Number", vbExclamation
        GoTo CleanExit
    End If
    
    If qcTested = "" Then
        MsgBox "Please complete QC Tested field", vbExclamation
        GoTo CleanExit
    End If
    
    ' Confirmation prompt
    Dim confirmMsg As String
    confirmMsg = "QC Number " & qcNumber & ", with Description """ & StandardName & """ will be released." & vbCrLf & vbCrLf & "Do you want to proceed?"
    
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "Confirm QC Release") = vbNo Then
        GoTo CleanExit
    End If
    
    Set mainWorkbook = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not fso.FileExists(filePath) Then
        MsgBox "QC Number database file not found: " & filePath, vbExclamation
        GoTo CleanExit
    End If
    
    ' *** GET NEXT QC_TRANS_ID BEFORE WRITING ANY FILES ***
    qcTransID = GetNextQCTransactionID()
    
    ' Read the entire file
    Set ts = fso.OpenTextFile(filePath, 1) ' ForReading
    fileContent = ts.ReadAll
    ts.Close
    
    lines = Split(fileContent, vbCrLf)
    found = False
    
    ' Get header indices
    Dim idxQC As Long
    Dim idxFirstRelDate As Long, idxLastRelDate As Long
    Dim idxFirstRelBy As Long, idxLastRelBy As Long
    Dim idxFirstRetestDate As Long, idxLastRetestDate As Long
    Dim idxFirstRetestBy As Long, idxLastRetestBy As Long
    Dim idxQCStatus As Long
    Dim idxTransDate As Long, idxTransUser As Long
    Dim idxNewRetest As Long
    Dim idxQCTested As Long, idxTestedDate As Long
    Dim idxChangeReason As Long
    Dim idxQCTransID As Long

    idxQC = GetHeaderIndex0(HEADER_QC_NUMBER)
    idxFirstRelDate = GetHeaderIndex0(HEADER_FIRST_RELEASE_DATE)
    idxLastRelDate = GetHeaderIndex0(HEADER_LAST_RELEASE_DATE)
    idxFirstRelBy = GetHeaderIndex0(HEADER_FIRST_RELEASE_BY)
    idxLastRelBy = GetHeaderIndex0(HEADER_LAST_RELEASE_BY)
    idxFirstRetestDate = GetHeaderIndex0(HEADER_FIRST_RETEST_DATE)
    idxLastRetestDate = GetHeaderIndex0(HEADER_LAST_RETEST_DATE)
    idxFirstRetestBy = GetHeaderIndex0(HEADER_FIRST_RETEST_BY)
    idxLastRetestBy = GetHeaderIndex0(HEADER_LAST_RETEST_BY)
    idxQCStatus = GetHeaderIndex0(HEADER_QC_STATUS_FLAG)
    idxTransDate = GetHeaderIndex0(HEADER_TRANS_DATE)
    idxTransUser = GetHeaderIndex0(HEADER_TRANS_USER)
    idxNewRetest = GetHeaderIndex0(HEADER_NEW_RETEST)
    idxQCTested = GetHeaderIndex0(HEADER_QC_TESTED)
    idxTestedDate = GetHeaderIndex0(HEADER_TESTED_DATE)
    idxChangeReason = GetHeaderIndex0(HEADER_CHANGE_REASON)
    idxQCTransID = GetHeaderIndex0(HEADER_QC_TRANS_ID)

    For i = 1 To UBound(lines)
        line = lines(i)
        If Trim(line) <> "" Then
            lineParts = Split(line, ",")
            If UBound(lineParts) >= idxQC And _
               LCase$(Trim$(Replace(lineParts(idxQC), """", ""))) = LCase$(Trim$(qcNumber)) Then

                ' Get the previous QC Status from the CSV
                previousQCStatus = Trim$(Replace(lineParts(idxQCStatus), """", ""))

                ' Check if current status is Rejected or Cancelled - cannot release
                If UCase$(Trim$(previousQCStatus)) = UCase$(QC_STATUS_REJECTED) Or _
                   UCase$(Trim$(previousQCStatus)) = UCase$(QC_STATUS_CANCELLED) Then
                    MsgBox "Cannot release QC Number " & qcNumber & "." & vbCrLf & vbCrLf & _
                           "Current status is '" & previousQCStatus & "' which is final and non-reversible." & vbCrLf & _
                           "No changes can be made to this record.", vbExclamation, "Release Not Allowed"
                    GoTo CleanExit
                End If

                ' Handle First/Last Release Date and By
                If Trim$(Replace(lineParts(idxFirstRelDate), """", "")) = "" Then
                    ' No first release - set both first and current
                    lineParts(idxFirstRelDate) = """" & DateToANSI(Format$(Now, "dd/mm/yyyy")) & """"
                    lineParts(idxFirstRelBy) = """" & Environ$("USERNAME") & """"
                Else
                    ' Already has first release - update last only
                    lineParts(idxLastRelDate) = """" & DateToANSI(Format$(Now, "dd/mm/yyyy")) & """"
                    lineParts(idxLastRelBy) = """" & Environ$("USERNAME") & """"
                End If
                
                ' Handle First/Last Retest Date and By
                HandleRetestDate lineParts, retestDate, idxFirstRetestDate, idxLastRetestDate, idxFirstRetestBy, idxLastRetestBy
                
                lineParts(idxQCStatus) = """Released"""
                lineParts(idxChangeReason) = """Released"""
                lineParts(idxTransDate) = """" & DateToANSI(Format$(Now, "dd/mm/yyyy")) & """"
                lineParts(idxTransUser) = """" & Environ$("USERNAME") & """"
                lineParts(idxQCTransID) = """" & qcTransID & """"
                
                ' Before accessing lineParts(idxNewRetest), add:
                If UBound(lineParts) < idxNewRetest Then
                    ReDim Preserve lineParts(idxNewRetest)
                End If
                lineParts(idxNewRetest) = """" & NewRetest & """"

                ' Before accessing lineParts(idxQCTested), add:
                If UBound(lineParts) < idxQCTested Then
                    ReDim Preserve lineParts(idxQCTested)
                End If
                lineParts(idxQCTested) = """" & qcTested & """"
                
                ' Before accessing lineParts(idxTestedDate), add:
                If UBound(lineParts) < idxTestedDate Then
                    ReDim Preserve lineParts(idxTestedDate)
                End If
                lineParts(idxTestedDate) = """" & DateToANSI(dateTested) & """"

                ' Update the line in the array
                lines(i) = Join(lineParts, ",")
                found = True
                Exit For
            End If
        End If
    Next i

    If Not found Then
        MsgBox "QC Number " & qcNumber & " not found in database.", vbExclamation
        GoTo CleanExit
    End If
    
    ' *** WRITE TO QC_NUMBER_CHANGES.CSV FIRST (AUDIT TRAIL) ***
    Set ts = fso.OpenTextFile(auditFilePath, 8, True) ' 8 = ForAppending
    ts.WriteLine lines(i)
    ts.Close
    Set ts = Nothing
    
    ' Write the updated content back to the file
    Set ts = fso.OpenTextFile(filePath, 2) ' ForWriting
    ts.Write Join(lines, vbCrLf)
    ts.Close
    
    MsgBox "QC Number " & qcNumber & " has been released successfully.", vbInformation

    ' Refresh the data
    Call LoadQCNumberData

    ' Clear contents
    Call ClearQCRelease
    
CleanExit:
    SafeRelease ts
    SafeRelease fso
    SafeRelease mainWorkbook
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Public Sub SaveQCStatus()
    
    ' Check for locked files first
    Dim qcFilePath As String, auditFilePath As String
    qcFilePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    auditFilePath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    
    If IsFileLocked(qcFilePath) Then
        MsgBox "Cannot proceed - QC Number file is in use by another user." & vbCrLf & _
               "Please try again when the file is available.", vbExclamation
        GoTo CleanExit
    End If
    
    If IsFileLocked(auditFilePath) Then
        MsgBox "Cannot proceed - QC Number Changes file is in use by another user." & vbCrLf & _
               "Please try again when the file is available.", vbExclamation
        GoTo CleanExit
    End If
    
    Dim mainWorkbook As Workbook
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim fileContent As String
    Dim lines() As String
    Dim line As String
    Dim lineParts() As String
    Dim found As Boolean
    Dim i As Long
    Dim qcTransID As String
    
    ' Declare updatable fields
    Dim qcNumber As String
    Dim StandardName As String
    Dim rmCode As String
    Dim VendorName As String
    Dim mfgName As String
    Dim MfgLot As String
    Dim MfgExpiry As String
    Dim numContainers As String
    Dim TotalQty As String
    Dim masterUoM As String
    Dim qcSampleQty As String
    Dim retentionSampleQty As String
    Dim QCBalance As String
    Dim nilStockFlag As String
    Dim QCStatus As String
    Dim qcTested As String
    Dim dateTested As String
    Dim retestDate As String
    Dim ChangeReason As String
    Dim NewRetest As String
    Dim previousQCStatus As String
    Dim statusChanged As Boolean
    Dim currentBalance As Double
    Dim greBalance As Double
    Dim samBalance As Double
    Dim totalBalance As Double
    Dim grrBalance As Double
    Dim sarBalance As Double
    Dim remainingBalance As Double
    Dim cancelValidationFailed As Boolean

    ' Get all field values that can be updated
    qcNumber = GetFieldValByLabelC("QC Number:")
    StandardName = GetFieldValByLabelC("Standard Name:")
    rmCode = GetFieldValByLabelC("RM Code:")
    VendorName = GetFieldValByLabelC("Vendor Name:")
    mfgName = GetFieldValByLabelC("Manufacturer Name:")
    MfgLot = GetFieldValByLabelC("Manufacturer Lot Number:")
    MfgExpiry = GetFieldValByLabelC("Manufacturer Expiry:")
    numContainers = GetFieldValByLabelC("No. Containers:")
    TotalQty = GetFieldValByLabelC("Total Quantity:")
    masterUoM = GetFieldValByLabelC("Unit of Measure:")
    qcSampleQty = GetFieldValByLabelC("QC Sample QTY:")
    retentionSampleQty = GetFieldValByLabelC("Retention Sample QTY:")
    QCBalance = GetFieldValByLabelC("QC Balance:")
    nilStockFlag = GetFieldValByLabelC("Nil Stock:")
    QCStatus = GetFieldValByLabelC("QC Status:")
    qcTested = GetFieldValByLabelC("QC Tested:")
    dateTested = GetFieldValByLabelC("Date Tested:")
    retestDate = GetFieldValByLabelC("Retest Date:")
    ChangeReason = GetFieldValByLabelC("Change Reason:")
    NewRetest = GetFieldValByLabelC("New or Retest:")

    If qcNumber = "" Then
        MsgBox "Please enter a QC Number", vbExclamation
        GoTo CleanExit
    End If
    
    Set mainWorkbook = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filePath = GetWorkbookTableFilePath(FILE_QC_NUMBER)
    
    If Not fso.FileExists(filePath) Then
        MsgBox "QC Number database file not found: " & filePath, vbExclamation
        GoTo CleanExit
    End If

    If ChangeReason = "" Then
        MsgBox "Please provide a Change Reason for the update.", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- Comma validation ---
     If Not ValidateNoCommas(ChangeReason, "Change Reason") Then GoTo CleanExit
    
    ' *** GET NEXT QC_TRANS_ID BEFORE ANY UPDATES ***
    qcTransID = GetNextQCTransactionID()
    
    ' Read the entire file
    Set ts = fso.OpenTextFile(filePath, 1) ' ForReading
    fileContent = ts.ReadAll
    ts.Close
    
    lines = Split(fileContent, vbCrLf)
    found = False
    
    ' Get header indices
    Dim idxQC As Long, idxStdName As Long, idxRMCode As Long
    Dim idxVendorName As Long, idxMfgName As Long, idxMfgLot As Long, idxMfgExpiry As Long
    Dim idxNumContainers As Long, idxTotalQty As Long, idxMasterUoM As Long
    Dim idxQCSample As Long, idxRetSample As Long, idxQCBalance As Long, idxQCStatus As Long
    Dim idxFirstRelDate As Long, idxLastRelDate As Long
    Dim idxFirstRelBy As Long, idxLastRelBy As Long
    Dim idxFirstRetestDate As Long, idxLastRetestDate As Long
    Dim idxFirstRetestBy As Long, idxLastRetestBy As Long
    Dim idxChangeReason As Long
    Dim idxTransDate As Long, idxTransUser As Long
    Dim idxNewRetest As Long, idxQCTested As Long
    Dim idxNilStockFlag As Long, idxNilStockDate As Long, idxNilStockBy As Long
    Dim idxTestedDate As Long
    Dim idxQCTransID As Long

    idxQC = GetHeaderIndex0(HEADER_QC_NUMBER)
    idxStdName = GetHeaderIndex0(HEADER_STANDARD_NAME)
    idxRMCode = GetHeaderIndex0(HEADER_RM_CODE)
    idxVendorName = GetHeaderIndex0(HEADER_VENDOR_NAME)
    idxMfgName = GetHeaderIndex0(HEADER_MANUFACTURER_NAME)
    idxMfgLot = GetHeaderIndex0(HEADER_MANUFACTURER_LOT)
    idxMfgExpiry = GetHeaderIndex0(HEADER_MANUFACTURER_EXPIRY)
    idxNumContainers = GetHeaderIndex0(HEADER_NUMBER_CONTAINERS)
    idxMasterUoM = GetHeaderIndex0(HEADER_MASTER_UOM)
    idxQCBalance = GetHeaderIndex0(HEADER_QC_BALANCE)
    idxQCStatus = GetHeaderIndex0(HEADER_QC_STATUS_FLAG)
    idxFirstRelDate = GetHeaderIndex0(HEADER_FIRST_RELEASE_DATE)
    idxLastRelDate = GetHeaderIndex0(HEADER_LAST_RELEASE_DATE)
    idxFirstRelBy = GetHeaderIndex0(HEADER_FIRST_RELEASE_BY)
    idxLastRelBy = GetHeaderIndex0(HEADER_LAST_RELEASE_BY)
    idxFirstRetestDate = GetHeaderIndex0(HEADER_FIRST_RETEST_DATE)
    idxLastRetestDate = GetHeaderIndex0(HEADER_LAST_RETEST_DATE)
    idxFirstRetestBy = GetHeaderIndex0(HEADER_FIRST_RETEST_BY)
    idxLastRetestBy = GetHeaderIndex0(HEADER_LAST_RETEST_BY)
    idxChangeReason = GetHeaderIndex0(HEADER_CHANGE_REASON)
    idxTransDate = GetHeaderIndex0(HEADER_TRANS_DATE)
    idxTransUser = GetHeaderIndex0(HEADER_TRANS_USER)
    idxNewRetest = GetHeaderIndex0(HEADER_NEW_RETEST)
    idxQCTested = GetHeaderIndex0(HEADER_QC_TESTED)
    idxTestedDate = GetHeaderIndex0(HEADER_TESTED_DATE)
    idxNilStockFlag = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG)
    idxNilStockBy = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_BY)
    idxNilStockDate = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_DATE)
    idxQCTransID = GetHeaderIndex0(HEADER_QC_TRANS_ID)

    For i = 1 To UBound(lines)
        line = lines(i)
        If Trim(line) <> "" Then
            lineParts = Split(line, ",")
            
            ' Add bounds checking before array access
            If UBound(lineParts) >= idxQC Then
            
                If LCase$(Trim$(Replace(lineParts(idxQC), """", ""))) = LCase$(Trim$(qcNumber)) Then
                    
                    ' Get the previous QC Status from the CSV for comparison
                    previousQCStatus = Trim$(Replace(lineParts(idxQCStatus), """", ""))

                    ' Check if previous status is Rejected or Cancelled - status is locked
                    If UCase$(Trim$(previousQCStatus)) = UCase$(QC_STATUS_REJECTED) Or _
                       UCase$(Trim$(previousQCStatus)) = UCase$(QC_STATUS_CANCELLED) Then
                        
                        ' Check if user is trying to change the status
                        If UCase$(Trim$(QCStatus)) <> UCase$(Trim$(previousQCStatus)) Then
                            MsgBox "Cannot change QC Status once it has been set to '" & previousQCStatus & "'." & vbCrLf & _
                                    "This status is final and non-reversible." & vbCrLf & vbCrLf & _
                                    "You may still update other fields associated with this record.", vbExclamation
                            GoTo CleanExit
                        End If
                    End If
                    
                    ' Determine if status is being changed
                    statusChanged = UCase$(Trim$(QCStatus)) <> UCase$(Trim$(previousQCStatus))
                    
                    ' Validate cancellation BEFORE any file writes
                    If statusChanged And UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_CANCELLED) Then
                        greBalance = GetQCTransactionBalance(qcNumber, TRANS_TYPE_GOODS_RECEIPT)
                        samBalance = GetQCTransactionBalance(qcNumber, TRANS_TYPE_SAMPLE)
                        
                        If greBalance = 0 Then
                            MsgBox "Cannot cancel QC Number " & qcNumber & ". No Goods Receipt (GRE) transaction found." & vbCrLf & _
                                    "Use 'Reject' status instead to write off this QC Number.", vbExclamation
                            cancelValidationFailed = True
                            GoTo CleanExit
                        End If
                        
                        totalBalance = GetQCBalanceFromRMTrans(qcNumber)
                        grrBalance = GetQCTransactionBalance(qcNumber, TRANS_TYPE_GOODS_RECEIPT_REVERSAL)
                        sarBalance = GetQCTransactionBalance(qcNumber, TRANS_TYPE_SAMPLING_REVERSAL)
                        remainingBalance = totalBalance - greBalance - samBalance
                        
                        If remainingBalance <> 0 Then
                            MsgBox "Cannot cancel QC Number " & qcNumber & ". This QC Number has consumption or other transactions." & vbCrLf & _
                                    "Only Goods Receipt (GRE) and Sampling (SAM) transactions are allowed for cancellation." & vbCrLf & _
                                    "Use 'Reject' status instead to write off this QC Number.", vbExclamation
                            cancelValidationFailed = True
                            GoTo CleanExit
                        End If
                    End If
                    
                    ' If changing to Rejected or Cancelled, show warning and require confirmation
                    If statusChanged Then
                        If UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_REJECTED) Or _
                           UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_CANCELLED) Then
                            
                            Dim warningMsg As String
                            Dim statusAction As String
                            
                            If UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_REJECTED) Then
                                statusAction = "REJECTED"
                            Else
                                statusAction = "CANCELLED"
                            End If
                            
                            warningMsg = "WARNING: This action is NON-REVERSIBLE!" & vbCrLf & vbCrLf & _
                                        "You are about to mark QC Number " & qcNumber & " as " & statusAction & "." & vbCrLf & vbCrLf & _
                                        "Once " & statusAction & ", this status cannot be changed." & vbCrLf & _
                                        "All associated inventory will be written off." & vbCrLf & vbCrLf & _
                                        "Do you want to continue?"
                            
                            Dim confirmResult As VbMsgBoxResult
                            confirmResult = MsgBox(warningMsg, vbExclamation + vbOKCancel, "Confirm Non-Reversible Action")
                            
                            If confirmResult <> vbOK Then
                                MsgBox "Update cancelled. No changes were made.", vbInformation
                                GoTo CleanExit
                            End If

                        End If
                    End If

                    ' Update basic fields
                    lineParts(idxStdName) = """" & StandardName & """"
                    lineParts(idxRMCode) = """" & rmCode & """"
                    lineParts(idxVendorName) = """" & VendorName & """"
                    lineParts(idxMfgName) = """" & mfgName & """"
                    lineParts(idxMfgLot) = """" & MfgLot & """"
                    lineParts(idxMfgExpiry) = """" & MfgExpiry & """"
                    lineParts(idxNumContainers) = """" & numContainers & """"
                    lineParts(idxMasterUoM) = """" & masterUoM & """"
                    lineParts(idxQCStatus) = """" & QCStatus & """"
                    lineParts(idxChangeReason) = """" & ChangeReason & """"
                    
                    ' Stamp transaction metadata
                    lineParts(idxTransDate) = """" & DateToANSI(Format$(Now, "dd/mm/yyyy")) & """"
                    lineParts(idxTransUser) = """" & Environ$("USERNAME") & """"
                    
                    ' Handle First/Last Retest Date and By
                    HandleRetestDate lineParts, retestDate, idxFirstRetestDate, idxLastRetestDate, idxFirstRetestBy, idxLastRetestBy
                    
                    ' Handle First/Last Release Date and By if status is "Released"
                    If UCase$(QCStatus) = "RELEASED" Then
                        ' Only update release dates if previous status was not already "Released"
                        If UCase$(Trim$(previousQCStatus)) <> "RELEASED" Then
                            If Trim$(Replace(lineParts(idxFirstRelDate), """", "")) = "" Then
                                ' No first release - set both first and current
                                lineParts(idxFirstRelDate) = """" & DateToANSI(Format$(Now, "dd/mm/yyyy")) & """"
                                lineParts(idxFirstRelBy) = """" & Environ$("USERNAME") & """"
                            Else
                                ' Already has first release - update last only
                                lineParts(idxLastRelDate) = """" & DateToANSI(Format$(Now, "dd/mm/yyyy")) & """"
                                lineParts(idxLastRelBy) = """" & Environ$("USERNAME") & """"
                            End If
                        End If
                    End If
                    
                    ' *** UPDATE QC_TRANS_ID ***
                    If idxQCTransID >= 0 And UBound(lineParts) >= idxQCTransID Then
                        lineParts(idxQCTransID) = """" & qcTransID & """"
                    End If
                    
                    ' Set Nil Stock Flag if status changed to Rejected or Cancelled
                    If statusChanged And (UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_REJECTED) Or UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_CANCELLED)) Then
                        lineParts(idxNilStockFlag) = """Y"""
                        lineParts(idxNilStockBy) = """" & Environ$("Username") & """"
                        lineParts(idxNilStockDate) = """" & DateToANSI(Format$(Date, "dd/mm/yyyy")) & """"

                        ' Set QC_Balance to zero for both Rejected and Cancelled
                        lineParts(idxQCBalance) = """0"""
                        
                    End If

                    ' Before accessing lineParts(idxQCTested), add:
                    If UBound(lineParts) < idxQCTested Then
                        ReDim Preserve lineParts(idxQCTested)
                    End If
                    lineParts(idxQCTested) = """" & qcTested & """"
                    
                    ' Before accessing lineParts(idxNewRetest), add:
                    If UBound(lineParts) < idxNewRetest Then
                        ReDim Preserve lineParts(idxNewRetest)
                    End If
                    lineParts(idxNewRetest) = """" & NewRetest & """"

                    ' Before accessing lineParts(idxTestedDate), add:
                    If UBound(lineParts) < idxTestedDate Then
                        ReDim Preserve lineParts(idxTestedDate)
                    End If
                    lineParts(idxTestedDate) = """" & DateToANSI(dateTested) & """"

                    ' Update the line in the array
                    lines(i) = Join(lineParts, ",")
                    found = True
                    Exit For
                End If
            End If
        End If
    Next i

    If Not found Then
        MsgBox "QC Number " & qcNumber & " not found in database.", vbExclamation
        GoTo CleanExit
    End If
    
    ' *** WRITE TO QC_NUMBER_CHANGES.CSV FIRST (AUDIT TRAIL) ***
    Dim auditPath As String
    auditPath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    Set ts = fso.OpenTextFile(auditPath, 8, True) ' 8 = ForAppending
    ts.WriteLine lines(i)
    ts.Close
    Set ts = Nothing
    
    ' Write the updated content back to the file
    Set ts = fso.OpenTextFile(filePath, 2) ' ForWriting
    ts.Write Join(lines, vbCrLf)
    ts.Close
    
    ' Check if QC Status changed to "Rejected"
    If statusChanged And UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_REJECTED) Then
        
        ' Get current balance from RM_Trans
        On Error GoTo WriteRejError
        currentBalance = GetQCBalanceFromRMTrans(qcNumber)
        
        If currentBalance > 0 Then
            ' Create REJ transaction to write off the remaining balance with QC_Trans_ID
            WriteRMTransaction TRANS_TYPE_REJECTED, qcNumber, QC_STATUS_REJECTED, _
                              -currentBalance, masterUoM, qcTransID, ChangeReason
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Check if QC Status changed to "Cancelled"
    If statusChanged And UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_CANCELLED) Then
        
        On Error GoTo WriteCancelError
        
        ' Create GRR (Goods Receipt Reversal) transaction if there are GRE transactions
        If greBalance <> 0 Then
            ' GRR reverses the GRE - use negative of the GRE balance with QC_Trans_ID
            WriteRMTransaction TRANS_TYPE_GOODS_RECEIPT_REVERSAL, qcNumber, QC_STATUS_CANCELLED, _
                              -greBalance, masterUoM, qcTransID, ChangeReason
        End If
        
        ' Create SAR (Sampling Reversal) transaction if there are SAM transactions
        If samBalance <> 0 Then
            ' SAR reverses the SAM - use negative of the SAM balance with QC_Trans_ID
            WriteRMTransaction TRANS_TYPE_SAMPLING_REVERSAL, qcNumber, QC_STATUS_CANCELLED, _
                              -samBalance, masterUoM, qcTransID, ChangeReason
        End If
        
        On Error GoTo ErrorHandler
    End If

    MsgBox "QC Number " & qcNumber & " has been updated successfully.", vbInformation
    
    ' Refresh the data
    Call LoadQCNumberData
    
    ' Clear contents
    Call ClearMaintainQCNumber
    
    GoTo CleanExit

WriteRejError:
    MsgBox "Error creating REJ transaction: " & Err.Description, vbCritical
    Resume Next

WriteCancelError:
    MsgBox "Error creating cancellation reversal transactions: " & Err.Description, vbCritical
    Resume Next

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
    
CleanExit:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set mainWorkbook = Nothing
    Exit Sub
End Sub

Private Sub HandleRetestDate(ByRef lineParts() As String, ByVal retestDate As String, _
                             ByVal idxFirstRetestDate As Long, ByVal idxLastRetestDate As Long, _
                             ByVal idxFirstRetestBy As Long, ByVal idxLastRetestBy As Long)
    If Trim$(retestDate) = "" Then
        ' Move last retest to first retest, then blank last retest fields
        lineParts(idxFirstRetestDate) = lineParts(idxLastRetestDate)
        lineParts(idxFirstRetestBy) = lineParts(idxLastRetestBy)
        lineParts(idxLastRetestDate) = """"""
        lineParts(idxLastRetestBy) = """"""
    Else
        ' Convert the input date directly to ANSI format
        Dim ansiDate As String
        ansiDate = DateToANSI(retestDate)
        
        If ansiDate <> "" Then
            ' Check if last retest date exists
            Dim lastRetestDate As String
            lastRetestDate = Trim$(Replace(lineParts(idxLastRetestDate), """", ""))
            
            If lastRetestDate = "" Then
                ' No last retest - set both first and last to the new retest date
                lineParts(idxFirstRetestDate) = """" & ansiDate & """"
                lineParts(idxFirstRetestBy) = """" & Environ$("USERNAME") & """"
                lineParts(idxLastRetestDate) = """" & ansiDate & """"
                lineParts(idxLastRetestBy) = """" & Environ$("USERNAME") & """"
            Else
                ' Check if new date matches existing last retest date
                If ansiDate = lastRetestDate Then
                    ' Same date - do nothing and exit
                    Exit Sub
                Else
                    ' Different date - move last to first and set new last
                    lineParts(idxFirstRetestDate) = lineParts(idxLastRetestDate)
                    lineParts(idxFirstRetestBy) = lineParts(idxLastRetestBy)
                    lineParts(idxLastRetestDate) = """" & ansiDate & """"
                    lineParts(idxLastRetestBy) = """" & Environ$("USERNAME") & """"
                End If
            End If
        End If
    End If
End Sub
