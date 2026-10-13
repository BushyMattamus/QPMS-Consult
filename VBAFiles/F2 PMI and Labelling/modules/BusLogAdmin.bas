Attribute VB_Name = "BusLogAdmin"
'Sub AdminOp()
'    Dim ws As Worksheet
'    Dim userPassword As String
'    Dim PW As String
'    Dim APW As String
'
'    ' Get General Password
'    PW = GetGeneralPassword()
'    APW = GetAdminPassword()
'
'    ' Prompt for password
'    userPassword = InputBox("Enter the password to proceed:", "Password Required")
'
'    ' Check if the password is correct
'    If userPassword <> APW Then
'        MsgBox "Incorrect password, operation cancelled.", vbCritical
'        Exit Sub
'    End If
'
'    ' Disable events and screen updating for performance
'    AppStateUtils.DisableAppUpdates
'
'    ' Enable "Show Sheet Tabs" in Display Options
'    ActiveWindow.DisplayWorkbookTabs = True
'
'    ' Loop through each worksheet
'    For Each ws In ThisWorkbook.Worksheets
'        ' Only unhide if it's NOT the Helper Lists sheet
'        If ws.Name <> "Helper Lists" Then
'            ws.Visible = xlSheetVisible
'            ws.Activate
'
'            ' Ensure row and column headers are shown
'            ActiveWindow.DisplayHeadings = True
'
'            ' Unprotect the worksheet
'            On Error Resume Next
'            ws.Unprotect password:=PW
'            On Error GoTo 0
'        End If
'    Next ws
'
'    Worksheets("Admin").Activate
'
'    ' Re-enable events and screen updating after operations
'    AppStateUtils.RestoreAppState
'
'    MsgBox "Operation completed successfully!", vbInformation
'End Sub
'
'Sub AdminEx()
'    Dim ws As Worksheet
'    Dim userPassword As String
'    Dim protectedSheets As Object
'    Dim mainMenuSheet As Worksheet
'    Dim sheetName As String
'    Dim PW As String
'    Dim APW As String
'
'    ' Get General Password
'    PW = GetGeneralPassword()
'    APW = GetAdminPassword()
'
'    ' Define the Main Menu
'    sheetName = "Main Menu"
'
'    ' Create a dictionary of sheets to protect and hide headers
'    Set protectedSheets = CreateObject("Scripting.Dictionary")
'    protectedSheets.Add "Label & PMI Print", True
'    protectedSheets.Add "Print Reprint", True
'    protectedSheets.Add "Label Register Review", True
'    protectedSheets.Add "Previous Prints", True
'    protectedSheets.Add "Label Maintenance", True
'    protectedSheets.Add "Avery Maintenance", True
'    protectedSheets.Add "Email Reports", True
'
'    ' Prompt for Admin Password
'    userPassword = InputBox("Enter the admin password to proceed:", "Admin Access Required")
'
'    ' Verify password
'    If userPassword <> APW Then
'        MsgBox "Incorrect password, operation cancelled.", vbCritical
'        Exit Sub
'    End If
'
'    ' Disable screen updating & events for better performance
'    AppStateUtils.DisableAppUpdates
'
'    ' Disable "Show Sheet Tabs" (Hides the tab bar at the bottom)
'    ActiveWindow.DisplayWorkbookTabs = False
'
'    ' Loop through each worksheet
'    For Each ws In ThisWorkbook.Worksheets
'        ' Check if the sheet should be protected
'        If protectedSheets.Exists(ws.Name) Then
'            ' Activate the worksheet
'            ws.Activate
'
'            ' Hide row and column headers
'            ActiveWindow.DisplayHeadings = False
'
'            ' Protect the worksheet
'            On Error Resume Next
'            ws.Protect password:=PW
'            On Error GoTo 0
'        Else
'            ' Activate the worksheet
'            ws.Activate
'
'            ' Ensure row and column headers are visible for unprotected sheets
'            ActiveWindow.DisplayHeadings = True
'
'            ' Unprotect the worksheet
'            On Error Resume Next
'            ws.Unprotect password:=PW
'            On Error GoTo 0
'        End If
'
'        ' Handle visibility
'        Select Case ws.Name
'            Case sheetName
'                ws.Visible = xlSheetVisible
'                Set mainMenuSheet = ws
'            Case "Helper Lists"
'                ws.Visible = xlSheetHidden   ' <<< make this sheet very hidden
'            Case Else
'                ws.Visible = xlSheetHidden
'        End Select
'
'    Next ws
'
'    ' Make sure "Main Menu" is visible and activated
'    If Not mainMenuSheet Is Nothing Then
'        mainMenuSheet.Visible = xlSheetVisible
'        mainMenuSheet.Activate
'        'mainMenuSheet.Protect password:=PW
'
'        ' Hide row and column headers for "Main Menu"
'        ActiveWindow.DisplayHeadings = False
'    Else
'        MsgBox "Error: 'Main Menu' sheet not found!", vbCritical
'    End If
'
'    ' Re-enable screen updating & events
'    AppStateUtils.RestoreAppState
'
'    MsgBox "Admin Exit completed. Only 'Main Menu' is visible, headers are hidden, and sheet tabs are disabled.", vbInformation
'End Sub
'