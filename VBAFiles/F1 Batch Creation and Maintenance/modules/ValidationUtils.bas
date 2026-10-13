Attribute VB_Name = "ValidationUtils"
' ========================================================================
' Validate that the expiry date matches the shelf life calculation
' ========================================================================
Public Function ValidateExpiryDate(ByVal mfgDate As String, ByVal expiryDate As String, _
                                     ByVal shelfNo As String, ByVal shelfLife As String, _
                                     ByRef userResponse As VbMsgBoxResult) As Boolean
    On Error GoTo ErrHandler

    Dim mfgDateValue As Date
    Dim expiryDateValue As Date
    Dim expectedExpiryDate As Date
    Dim shelfNum As Long
    Dim shelfUnit As String
    Dim calculatedDiff As Long
    Dim actualDiff As Long
    Dim warningMsg As String

    ValidateExpiryDate = True
    userResponse = vbYes

    ' Skip validation if shelf life is "AS PER MANUFACTURER"
    shelfUnit = UCase(Trim(shelfLife))
    If shelfUnit = "AS PER MANUFACTURER" Then
        GoTo CleanExit
    End If

    ' Validate inputs
    If Not IsDate(mfgDate) Or Not IsDate(expiryDate) Then
        GoTo CleanExit
    End If

    mfgDateValue = CDate(mfgDate)
    expiryDateValue = CDate(expiryDate)
    shelfNum = val(shelfNo)

    If shelfNum <= 0 Then
        GoTo CleanExit
    End If

    ' Calculate expected expiry date based on shelf life
    Select Case shelfUnit
        Case "DAYS", "DAY"
            expectedExpiryDate = DateAdd("d", shelfNum, mfgDateValue)
        Case "MONTHS", "MONTH"
            expectedExpiryDate = DateAdd("m", shelfNum, mfgDateValue)
        Case "YEARS", "YEAR"
            expectedExpiryDate = DateAdd("yyyy", shelfNum, mfgDateValue)
        Case Else
            GoTo CleanExit
    End Select

    ' Compare expected vs actual expiry date
    If expectedExpiryDate <> expiryDateValue Then
        warningMsg = "Warning: Expiry date validation mismatch!" & vbCrLf & vbCrLf & _
                     "Manufacturing Date: " & Format(mfgDateValue, "dd/mm/yyyy") & vbCrLf & _
                     "Shelf Life: " & shelfNo & " " & shelfLife & vbCrLf & _
                     "Expected Expiry Date: " & Format(expectedExpiryDate, "dd/mm/yyyy") & vbCrLf & _
                     "Actual Expiry Date: " & Format(expiryDateValue, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
                     "Do you want to proceed with this batch?"

        userResponse = MsgBox(warningMsg, vbExclamation + vbYesNo, "Expiry Date Mismatch")

        If userResponse = vbNo Then
            ValidateExpiryDate = False
        End If
    End If

    GoTo CleanExit

ErrHandler:
    ValidateExpiryDate = True
    userResponse = vbYes
    GoTo CleanExit

CleanExit:
    Exit Function
End Function
