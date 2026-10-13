Attribute VB_Name = "BusLogMaintainQCNumber"
Option Explicit

Public Sub SaveMaintainQCNumber()
    DisableAppUpdates
    
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
    Dim poNumber As String
    Dim vendorName As String
    Dim mfgName As String
    Dim mfgLot As String
    Dim mfgExpiry As String
    Dim deliveryQty As String
    Dim numContainers As String
    Dim qcBalance As String
    Dim masterUoM As String
    Dim QCStatus As String
    Dim changeReason As String
    Dim itemCostStr As String
    Dim itemCost As Double
    Dim previousQCStatus As String
    Dim currentBalance As Double
    Dim greBalance As Double
    Dim samBalance As Double
    Dim totalBalance As Double
    Dim remainingBalance As Double
    Dim statusChanged As Boolean
    Dim qtyForCalc As Double
    Dim deliveryQtyStr As String
    Dim previousItemCost As String
    Dim basisLabel As String
    Dim costChangeMsg As String

    ' Get all field values that can be updated
    qcNumber = GetFieldValByLabelC("QC Number:")
    StandardName = GetFieldValByLabelC("Standard Name:")
    rmCode = GetFieldValByLabelC("RM Code:")
    vendorName = GetFieldValByLabelC("Vendor Name:")
    poNumber = GetFieldValByLabelC("PO Number:")
    mfgName = GetFieldValByLabelC("Manufacturer Name:")
    mfgLot = GetFieldValByLabelC("Manufacturer Lot Number:")
    mfgExpiry = GetFieldValByLabelC("Manufacturer Expiry:")
    numContainers = GetFieldValByLabelC("No. Containers:")
    deliveryQty = GetFieldValByLabelC("Delivery Qty:")
    qcBalance = GetFieldValByLabelC("QC Balance:")
    masterUoM = GetFieldValByLabelC("Unit of Measure:")
    QCStatus = GetFieldValByLabelC("QC Status:")
    changeReason = GetFieldValByLabelC("Change Reason:")
    itemCostStr = Trim$(GetFieldValByLabelC("Total Item Cost (excl GST):"))
    
    ' Validate Item Cost is numeric if provided
    If itemCostStr <> "" Then
        If Not IsNumeric(itemCostStr) Then
            MsgBox "Total Item Cost (excl GST) must be a valid number.", vbCritical
            GoTo CleanExit
        End If
        itemCost = CDbl(itemCostStr)
    Else
        itemCost = 0
    End If
    
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

    If changeReason = "" Then
        MsgBox "Please provide a Change Reason for the update.", vbExclamation
        GoTo CleanExit
    End If

    If Not ValidateNoCommas(changeReason, "Change Reason") Then
        GoTo CleanExit
    End If

    ' --- Comma validation for fields being written to file ---
    If Not ValidateNoCommas(poNumber, "PO Number") Then GoTo CleanExit
    If Not ValidateNoCommas(mfgLot, "Manufacturer Lot Number") Then GoTo CleanExit
    
    ' Read the entire file
    Set ts = fso.OpenTextFile(filePath, 1) ' ForReading
    fileContent = ts.ReadAll
    ts.Close
    
    lines = Split(fileContent, vbCrLf)
    found = False
    
    Dim idxQC As Long, idxStdName As Long, idxRMCode As Long
    Dim idxVendorName As Long, idxMfgName As Long, idxMfgLot As Long, idxMfgExpiry As Long
    Dim idxNumContainers As Long
    Dim idxMasterUoM As Long, idxQCBalance As Long
    Dim idxQCStatus As Long
    Dim idxChangeReason As Long
    Dim idxFirstRelDate As Long, idxLastRelDate As Long
    Dim idxFirstRelBy As Long, idxLastRelBy As Long
    Dim idxNilStockFlag As Long
    Dim idxNilStockBy As Long
    Dim idxNilStockDate As Long
    Dim idxTransDate As Long
    Dim idxTransUser As Long
    Dim idxQCTransID As Long
    Dim idxQCSample As Long
    Dim idxRetSample As Long
    Dim idxTotSample As Long
    Dim idxItemCost As Long
    Dim idxBuomCost As Long
    Dim idxdeliveryQty As Long
    Dim idxPONumber As Long
    
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
    idxChangeReason = GetHeaderIndex0(HEADER_CHANGE_REASON)
    idxFirstRelDate = GetHeaderIndex0(HEADER_FIRST_RELEASE_DATE)
    idxLastRelDate = GetHeaderIndex0(HEADER_LAST_RELEASE_DATE)
    idxFirstRelBy = GetHeaderIndex0(HEADER_FIRST_RELEASE_BY)
    idxLastRelBy = GetHeaderIndex0(HEADER_LAST_RELEASE_BY)
    idxNilStockFlag = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG)
    idxNilStockBy = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_BY)
    idxNilStockDate = GetHeaderIndex0(HEADER_NIL_STOCK_FLAG_DATE)
    idxTransDate = GetHeaderIndex0(HEADER_TRANS_DATE)
    idxTransUser = GetHeaderIndex0(HEADER_TRANS_USER)
    idxQCTransID = GetHeaderIndex0(HEADER_QC_TRANS_ID)
    idxQCSample = GetHeaderIndex0(HEADER_QC_SAMPLE_QTY)
    idxRetSample = GetHeaderIndex0(HEADER_RETENTION_SAMPLE_QTY)
    idxTotSample = GetHeaderIndex0(HEADER_TOTAL_SAMPLE_QTY)
    idxItemCost = GetHeaderIndex0(HEADER_ITEM_COST)
    idxBuomCost = GetHeaderIndex0(HEADER_BUOM_COST)
    idxdeliveryQty = GetHeaderIndex0(HEADER_TOTAL_QTY)
    idxPONumber = GetHeaderIndex0(HEADER_PO_NUMBER)
    ' *** GET NEXT QC_TRANS_ID BEFORE ANY UPDATES ***
    qcTransID = GetNextQCTransactionID()
    
    For i = 1 To UBound(lines)
        line = lines(i)
        If Trim(line) <> "" Then
            lineParts = Split(line, ",")

            ' Pad lineParts to match expected header count
            Dim expectedCols As Long
            expectedCols = UBound(GetQCNumberHeaders()) - LBound(GetQCNumberHeaders())
            If UBound(lineParts) < expectedCols Then
                ReDim Preserve lineParts(expectedCols)
                Dim p As Long
                For p = UBound(lineParts) To expectedCols
                    If lineParts(p) = "" Then lineParts(p) = """"""
                Next p
            End If
            
            If UBound(lineParts) >= idxQC And _
               LCase$(Trim$(Replace(lineParts(idxQC), """", ""))) = LCase$(Trim$(qcNumber)) Then
                
                ' Store previous QC status before updating
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
                        GoTo CleanExit
                    End If
                    
                    totalBalance = GetQCBalanceFromRMTrans(qcNumber)
                    remainingBalance = totalBalance - greBalance - samBalance
                    
                    If remainingBalance <> 0 Then
                        MsgBox "Cannot cancel QC Number " & qcNumber & ". This QC Number has consumption or other transactions." & vbCrLf & _
                                "Only Goods Receipt (GRE) and Sampling (SAM) transactions are allowed for cancellation." & vbCrLf & _
                                "Use 'Reject' status instead to write off this QC Number.", vbExclamation
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
                
                ' Capture previous values BEFORE updating lineParts
                Dim previousDeliveryQty As String
                Dim deliveryQtyChanged As Boolean
                Dim itemCostChanged As Boolean
                
                previousDeliveryQty = Trim$(Replace(lineParts(idxdeliveryQty), """", ""))
                previousItemCost = Trim$(Replace(lineParts(idxItemCost), """", ""))
                
                deliveryQtyChanged = (previousDeliveryQty <> deliveryQty)
                If itemCostStr <> "" Then
                    itemCostChanged = (previousItemCost <> itemCostStr)
                Else
                    itemCostChanged = False
                End If
                
                ' Update lineParts with new values
                lineParts(idxStdName) = """" & StandardName & """"
                lineParts(idxRMCode) = """" & rmCode & """"
                lineParts(idxVendorName) = """" & vendorName & """"
                lineParts(idxMfgName) = """" & mfgName & """"
                lineParts(idxMfgLot) = """" & mfgLot & """"
                lineParts(idxMfgExpiry) = """" & DateToANSI(mfgExpiry) & """"
                lineParts(idxNumContainers) = """" & numContainers & """"
                lineParts(idxdeliveryQty) = """" & deliveryQty & """"
                lineParts(idxMasterUoM) = """" & masterUoM & """"
                lineParts(idxQCStatus) = """" & QCStatus & """"
                lineParts(idxChangeReason) = """" & changeReason & """"

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
                
                ' Handle Rejected or Cancelled status changes
                If statusChanged And (UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_REJECTED) Or UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_CANCELLED)) Then
                    ' Set Nil Stock flags
                    lineParts(idxNilStockFlag) = """Y"""
                    lineParts(idxNilStockBy) = """" & Environ("Username") & """"
                    lineParts(idxNilStockDate) = """" & DateToANSI(Format(Date, "dd/mm/yyyy")) & """"
                    
                    ' Set QC_Balance to zero for both Rejected and Cancelled
                    lineParts(idxQCBalance) = """0"""
                    
                    ' For Cancelled only: reset sample quantities to zero
                    If UCase$(Trim$(QCStatus)) = UCase$(QC_STATUS_CANCELLED) Then
                        lineParts(idxQCSample) = """0"""
                        lineParts(idxRetSample) = """0"""
                        lineParts(idxTotSample) = """0"""
                    End If
                End If
                
                ' Handle Item Cost and BUOM Cost updates
                If itemCostStr <> "" Then
                    ' Recalculate BUoM Cost if either Item Cost or Delivery Qty changed
                    If itemCostChanged Or deliveryQtyChanged Then
                        deliveryQtyStr = deliveryQty
                        
                        If deliveryQtyStr <> "" And IsNumeric(deliveryQtyStr) Then
                            ' Use the (possibly updated) deliveryQty
                        Else
                            deliveryQtyStr = "0"
                        End If
                        
                        If CDbl(deliveryQtyStr) > 0 Then
                            qtyForCalc = CDbl(deliveryQtyStr)
                            basisLabel = "Delivered Quantity"
                        Else
                            qtyForCalc = CDbl(Val(Replace(lineParts(idxQCBalance), """", "")))
                            basisLabel = "QC Balance"
                        End If
                        
                        If qtyForCalc > 0 Then
                            Dim buomCost As Double
                            buomCost = Round(itemCost / qtyForCalc, 4)
                            
                            costChangeMsg = "Change in Item Cost or Delivery Quantity" & vbCrLf & vbCrLf & _
                                          "Item Cost: $" & Format(itemCost, "#,##0.00") & vbCrLf & _
                                          "Calculation Basis: " & basisLabel & " (" & Format(qtyForCalc, "#,##0.00") & " " & masterUoM & ")" & vbCrLf & _
                                          "Calculated BUoM Cost: $" & Format(buomCost, "#,##0.0000") & vbCrLf & vbCrLf & _
                                          "Do you want to proceed with this cost update?"
                            
                            Dim costConfirm As VbMsgBoxResult
                            costConfirm = MsgBox(costChangeMsg, vbQuestion + vbOKCancel, "Confirm Cost Change")
                            
                            If costConfirm <> vbOK Then
                                MsgBox "Cost update cancelled. No changes were made.", vbInformation
                                GoTo CleanExit
                            End If
                            
                            lineParts(idxItemCost) = """" & itemCost & """"
                            lineParts(idxBuomCost) = """" & buomCost & """"
                        End If
                    Else
                        lineParts(idxItemCost) = """" & itemCost & """"
                    End If
                End If

                ' Always record transaction metadata (current date + user)
                If idxTransDate >= 0 And UBound(lineParts) >= idxTransDate Then
                    lineParts(idxTransDate) = """" & DateToANSI(Format(Date, "dd/mm/yyyy")) & """"
                End If
                If idxTransUser >= 0 And UBound(lineParts) >= idxTransUser Then
                    lineParts(idxTransUser) = """" & Environ("Username") & """"
                End If
                
                ' *** UPDATE QC_TRANS_ID ***
                If idxQCTransID >= 0 And UBound(lineParts) >= idxQCTransID Then
                    lineParts(idxQCTransID) = """" & qcTransID & """"
                End If

                If idxPONumber >= 0 And UBound(lineParts) >= idxPONumber Then
                    lineParts(idxPONumber) = """" & poNumber & """"
                End If
                
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
    Dim auditPath As String
    auditPath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    Set ts = fso.OpenTextFile(auditPath, 8, True) ' 8 = ForAppending
    ts.WriteLine lines(i)
    ts.Close
    Set ts = Nothing
    
    ' Write the updated content back to QC_Number.csv
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
                              -currentBalance, masterUoM, qcTransID, changeReason
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
                              -greBalance, masterUoM, qcTransID, changeReason
        End If
        
        ' Create SAR (Sampling Reversal) transaction if there are SAM transactions
        If samBalance <> 0 Then
            ' SAR reverses the SAM - use negative of the SAM balance with QC_Trans_ID
            WriteRMTransaction TRANS_TYPE_SAMPLING_REVERSAL, qcNumber, QC_STATUS_CANCELLED, _
                              -samBalance, masterUoM, qcTransID, changeReason
        End If
        
        On Error GoTo ErrorHandler
    End If

    MsgBox "QC Number " & qcNumber & " has been updated successfully.", vbInformation
    
    ' Refresh the data
    Call LoadQCNumberData
    Call LoadRMMasterData
    
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
    RestoreAppState
    Exit Sub
End Sub
