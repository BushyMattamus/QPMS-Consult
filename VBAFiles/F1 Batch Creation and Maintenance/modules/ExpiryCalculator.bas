Attribute VB_Name = "ExpiryCalculator"
Public Sub CalculateExpiryDate(sheetName As String)

    On Error GoTo ErrHandler

    Dim mfgDateCell As Range
    Dim shelfUnitCell As Range
    Dim shelfNumCell As Range
    Dim expiryCell As Range

    Dim mfgDate As Variant
    Dim shelfUnit As String
    Dim shelfNum As Long
    Dim expiryDate As Date

    ' Get all the fields dynamically using label search
    Set mfgDateCell = GetFieldByLabel(sheetName, "C", "Manufacturing date")
    Set shelfUnitCell = GetFieldByLabel(sheetName, "C", "Shelf life(Days/Months/Years)")
    Set shelfNumCell = GetFieldByLabel(sheetName, "C", "Shelf life(No of days/Months/Years)")
    Set expiryCell = GetFieldByLabel(sheetName, "C", "Recommended expiry date")

    ' Read values
    mfgDate = mfgDateCell.value
    shelfUnit = Trim(UCase(shelfUnitCell.value))
    shelfNum = val(shelfNumCell.value)

    ' If no mfg date or no shelf life data ? clear expiry date
    If (mfgDate = "" Or Not IsDate(mfgDate)) _
        Or shelfNum <= 0 _
        Or shelfUnit = "" Then
        
        expiryCell.ClearContents
        Exit Sub
    End If

    ' Calculate expiry date based on unit
    Select Case shelfUnit
        Case "DAYS", "DAY"
            expiryDate = DateAdd("d", shelfNum, CDate(mfgDate))

        Case "MONTHS", "MONTH"
            expiryDate = DateAdd("m", shelfNum, CDate(mfgDate))

        Case "YEARS", "YEAR"
            expiryDate = DateAdd("yyyy", shelfNum, CDate(mfgDate))

        Case Else
            MsgBox "Invalid shelf life unit: " & shelfUnit, vbCritical
            Exit Sub
    End Select

    ' Write expiry date
    expiryCell.value = expiryDate
    expiryCell.NumberFormat = "dd/mm/yyyy"

    Exit Sub

ErrHandler:
    MsgBox "Error in CalculateExpiryDate:" & vbCrLf & Err.Description, vbCritical

End Sub

