Attribute VB_Name = "Navigation"
'Sub NavAveMaint()
'    ' Unhides and navigates to the "Avery Maintenance" sheet while hiding the current sheet
'
'    Dim ws As Worksheet
'    Dim wsCurrent As Worksheet
'    Dim password As String
'    Dim userInput As String
'    Dim LMPW As String
'    Dim PW As String
'
'    PW = GetGeneralPassword()
'    LMPW = GetLabelMaintenancePassword()
'
'
'    On Error Resume Next ' Handle error if the sheet does not exist
'    Set wsCurrent = ActiveSheet
'    Set ws = Worksheets("Avery Maintenance")
'
'    ' Define the correct password
'    password = LMPW
'
'    ' Prompt the user for the password
'    userInput = InputBox("Enter the password to access the sheet:", "Password Required")
'
'    ' Check if the password is correct
'    If userInput = password Then
'
'        MsgBox "Access granted.", vbInformation, "Success"
'
'        ' Force screen update to prevent pointer desync
'        Application.ScreenUpdating = False
'
'        ' Unhide Account Maintenance if necessary
'        ws.Visible = xlSheetVisible
'
'        ' Hide the current sheet
'        wsCurrent.Visible = xlSheetHidden
'
'        ' Navigate to the desired sheet
'        ws.Activate
'
'        ' Force a scroll reset to ensure proper alignment
'        ActiveWindow.ScrollRow = 1
'        ActiveWindow.ScrollColumn = 1
'
'        ' Enable screen updating
'        Application.ScreenUpdating = True
'
'    Else
'        ' Display an error message
'        MsgBox "Incorrect password. Access denied.", vbCritical, "Error"
'    End If
'
'Cleanup:
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'    On Error GoTo 0 ' Reset error handling
'End Sub
'
'Sub NavBatchEntry()
'    ' Unhides and navigates to the "Label & PMI Print" sheet while hiding the current sheet
'
'    Dim wsCurrent As Worksheet
'    Dim wsTarget As Worksheet
'
'    ' Set the current and target sheets
'    Set wsCurrent = ActiveSheet
'    Set wsTarget = ThisWorkbook.Sheets("Label & PMI Print")
'
'    ' If already on Target Sheet, do nothing
'    If wsCurrent.Name = wsTarget.Name Then GoTo Cleanup
'
'    ' Ensure the target sheet is visible
'    wsTarget.Visible = xlSheetVisible
'
'    ' Activate the target sheet
'    wsTarget.Activate
'
'    ' Hide the current sheet
'    wsCurrent.Visible = xlSheetHidden
'
'Cleanup:
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'End Sub
'
'Sub NavEmailReports()
'    ' Unhides and navigates to the "Email Reports" sheet while hiding the current sheet
'
'    Dim wsCurrent As Worksheet
'    Dim wsTarget As Worksheet
'
'    ' Set the current and target sheets
'    Set wsCurrent = ActiveSheet
'    Set wsTarget = ThisWorkbook.Sheets("Email Reports")
'
'    ' If already on Target Sheet, do nothing
'    If wsCurrent.Name = wsTarget.Name Then GoTo Cleanup
'
'    ' Ensure the target sheet is visible
'    wsTarget.Visible = xlSheetVisible
'
'    ' Activate the target sheet
'    wsTarget.Activate
'
'    ' Enter default dates in cells C4 and C5
'    wsTarget.Range("C4").value = Date - 7
'    wsTarget.Range("C5").value = Date
'
'    ' Hide the current sheet
'    wsCurrent.Visible = xlSheetHidden
'
'Cleanup:
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'End Sub
'
'Sub NavLabMaint()
'    ' Unhides and navigates to the "Label Maintenance" sheet while hiding the current sheet
'
'    Dim ws As Worksheet
'    Dim wsCurrent As Worksheet
'    Dim password As String
'    Dim userInput As String
'    Dim LMPW As String
'    Dim PW As String
'
'    PW = GetGeneralPassword()
'    LMPW = GetLabelMaintenancePassword()
'
'
'    On Error Resume Next ' Handle error if the sheet does not exist
'    Set wsCurrent = ActiveSheet
'    Set ws = Worksheets("Label Maintenance")
'
'    ' Define the correct password
'    password = LMPW
'
'    ' Prompt the user for the password
'    userInput = InputBox("Enter the password to access the sheet:", "Password Required")
'
'    ' Check if the password is correct
'    If userInput = password Then
'
'        MsgBox "Access granted.", vbInformation, "Success"
'
'        ' Force screen update to prevent pointer desync
'        Application.ScreenUpdating = False
'
'        ' Unhide Account Maintenance if necessary
'        ws.Visible = xlSheetVisible
'
'        ' Hide the current sheet
'        wsCurrent.Visible = xlSheetHidden
'
'        ' Navigate to the desired sheet
'        ws.Activate
'
'        ' Force a scroll reset to ensure proper alignment
'        ActiveWindow.ScrollRow = 1
'        ActiveWindow.ScrollColumn = 1
'
'        ' Enable screen updating
'        Application.ScreenUpdating = True
'
'    Else
'        ' Display an error message
'        MsgBox "Incorrect password. Access denied.", vbCritical, "Error"
'    End If
'
'Cleanup:
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'    On Error GoTo 0 ' Reset error handling
'End Sub
'
'Sub NavLabRegRev()
'    ' Unhides and navigates to the "Label Register Review" sheet while hiding the current sheet
'
'    Dim wsCurrent As Worksheet
'    Dim wsTarget As Worksheet
'
'    ' Disable events and screen updating
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'
'    ' Set the current and target sheets
'    Set wsCurrent = ActiveSheet
'    Set wsTarget = ThisWorkbook.Sheets("Label Register Review")
'
'    ' If already on Target Sheet, do nothing
'    If wsCurrent.Name = wsTarget.Name Then GoTo Cleanup
'
'    ' Ensure the target sheet is visible
'    wsTarget.Visible = xlSheetVisible
'
'    ' Activate the target sheet
'    wsTarget.Activate
'
'    ' Trigger Refresh for page
'    wsTarget.Range("C5").value = ActiveSheet.Range("C5").value
'
'    ' Hide the current sheet
'    wsCurrent.Visible = xlSheetHidden
'
'Cleanup:
'
'    ' Enable events and screen updating
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'End Sub
'
'Sub NavMainMenu()
'    ' Unhides and navigates to the "Main Menu" sheet while hiding the current sheet
'
'    Dim wsCurrent As Worksheet
'    Dim wsTarget As Worksheet
'
'    ' Set the current and target sheets
'    Set wsCurrent = ActiveSheet
'    Set wsTarget = ThisWorkbook.Sheets("Main Menu")
'
'    ' If already on Target Sheet, do nothing
'    If wsCurrent.Name = wsTarget.Name Then GoTo Cleanup
'
'    ' Ensure the target sheet is visible
'    wsTarget.Visible = xlSheetVisible
'
'    ' Activate the target sheet
'    wsTarget.Activate
'
'    ' Hide the current sheet
'    wsCurrent.Visible = xlSheetHidden
'
'Cleanup:
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'End Sub
'
'Sub NavPrevPrint()
'    ' Unhides and navigates to the "Previous Prints" sheet while hiding the current sheet
'
'    Dim wsCurrent As Worksheet
'    Dim wsTarget As Worksheet
'
'    ' Disable events and screen updating
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'
'    ' Set the current and target sheets
'    Set wsCurrent = ActiveSheet
'    Set wsTarget = ThisWorkbook.Sheets("Previous Prints")
'
'    ' If already on Target Sheet, do nothing
'    If wsCurrent.Name = wsTarget.Name Then GoTo Cleanup
'
'    ' Ensure the target sheet is visible
'    wsTarget.Visible = xlSheetVisible
'
'    ' Activate the target sheet
'    wsTarget.Activate
'
'    ' Trigger Refresh for page
'    wsTarget.Range("D5").value = ActiveSheet.Range("D5").value
'
'    ' Hide the current sheet
'    wsCurrent.Visible = xlSheetHidden
'
'Cleanup:
'
'    ' Enable events and screen updating
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'End Sub
'
'Sub NavPrntReprnt()
'    ' Unhides and navigates to the "Print Reprint" sheet while hiding the current sheet
'
'    Dim wsCurrent As Worksheet
'    Dim wsTarget As Worksheet
'
'    ' Set the current and target sheets
'    Set wsCurrent = ActiveSheet
'    Set wsTarget = ThisWorkbook.Sheets("Reprint Labels")
'
'    ' If already on Target Sheet, do nothing
'    If wsCurrent.Name = wsTarget.Name Then GoTo Cleanup
'
'    ' Ensure the target sheet is visible
'    wsTarget.Visible = xlSheetVisible
'
'    ' Activate the target sheet
'    wsTarget.Activate
'
'    ' Trigger Refresh for page
'    wsTarget.Range("C5").value = ActiveSheet.Range("C5").value
'
'    ' Hide the current sheet
'    wsCurrent.Visible = xlSheetHidden
'
'Cleanup:
'    Set wsCurrent = Nothing
'    Set wsTarget = Nothing
'End Sub
'
'
'
'