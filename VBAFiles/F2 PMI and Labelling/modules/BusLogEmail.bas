Attribute VB_Name = "BusLogEmail"
Sub GenerateSummaryEmail_Silent_DeleteSent()
    ' Worksheet variables
    Dim wsEmail As Worksheet
    Dim wsLabelPrints As Worksheet

    ' ListObject variables
    Dim tblEmailAdd As ListObject
    Dim tblEmailHistory As ListObject
    Dim tblLabelPrints As ListObject

    ' Outlook / General object variables
    Dim delFolder As Object
    Dim delItems As Object
    Dim filteredDelItems As Object
    Dim foundItem As Object
    Dim itm As Object
    Dim mailItem As Object
    Dim ns As Object
    Dim outlookApp As Object

    ' Variant variables
    Dim dataArray As Variant
    Dim filteredArray() As Variant
    Dim tempVal As Variant

    ' String variables
    Dim emailBody As String
    Dim filter As String
    Dim mailSubject As String
    Dim recipientList As String

    ' Date variables
    Dim lastEmailDate As Date
    Dim nowTime As Date

    ' Long variables
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lastRow As Long
    Dim rowCount As Long

    ' ListRow variables
    Dim newRow As ListRow
    Dim r As ListRow

    On Error GoTo Cleanup

    '--- Identify worksheets and tables ---
    Set wsEmail = ThisWorkbook.Sheets("Email")
    Set wsLabelPrints = ThisWorkbook.Sheets("Label Prints")

    Set tblEmailAdd = wsEmail.ListObjects("tblEmailAdd")
    Set tblLabelPrints = wsLabelPrints.ListObjects("tblLabelPrints")
    Set tblEmailHistory = wsEmail.ListObjects("tblEmailHistory")

    '--- Build recipient list ---
    recipientList = ""
    For Each r In tblEmailAdd.ListRows
        If LCase(r.Range(3).value) = "y" And LCase(r.Range(4).value) = "x" Then
            If recipientList <> "" Then
                recipientList = recipientList & "; "
            End If
            recipientList = recipientList & r.Range(2).value
        End If
    Next r
    If recipientList = "" Then GoTo Cleanup

    '--- Determine last email sent date/time from tblEmailHistory ---
    lastEmailDate = 0
    For Each r In tblEmailHistory.ListRows
        If r.Range(3).value = "Printed Report" Then
            If r.Range(2).value > lastEmailDate Then
                lastEmailDate = r.Range(2).value
            End If
        End If
    Next r

    '--- Extract data from tblLabelPrints ---
    If tblLabelPrints.DataBodyRange Is Nothing Then GoTo Cleanup
    lastRow = tblLabelPrints.DataBodyRange.Rows.Count
    If lastRow = 0 Then GoTo Cleanup

    dataArray = tblLabelPrints.DataBodyRange.value

    '--- Filter by Print Date in column 18 (since last email sent date/time) ---
    ReDim filteredArray(1 To UBound(dataArray, 1), 1 To UBound(dataArray, 2))
    rowCount = 0
    For i = 1 To UBound(dataArray, 1)
        If IsDate(dataArray(i, 18)) Then
            If dataArray(i, 18) >= lastEmailDate Then
                rowCount = rowCount + 1
                For j = 1 To UBound(dataArray, 2)
                    filteredArray(rowCount, j) = dataArray(i, j)
                Next j
            End If
        End If
    Next i
    If rowCount = 0 Then GoTo Cleanup

    '--- (Optional) Sort by Created Date (column 8), ascending ---
    If rowCount > 1 Then
        For i = 1 To rowCount - 1
            For j = i + 1 To rowCount
                If filteredArray(i, 8) > filteredArray(j, 8) Then
                    For k = 1 To UBound(filteredArray, 2)
                        tempVal = filteredArray(i, k)
                        filteredArray(i, k) = filteredArray(j, k)
                        filteredArray(j, k) = tempVal
                    Next k
                End If
            Next j
        Next i
    End If

    '--- Build the HTML table for the email body ---
    emailBody = "<p>Dear Team,</p>" & _
                "<p>Please find below the Label Print Summary report since " & lastEmailDate & ".</p>" & _
                "<table border='1' cellspacing='0' cellpadding='5'>" & _
                "<tr>" & _
                    "<th>QID Code</th>" & _
                    "<th>Product Description</th>" & _
                    "<th>Batch Number</th>" & _
                    "<th>Expiry</th>" & _
                    "<th>Created Date</th>" & _
                    "<th>Print Date</th>" & _
                    "<th>Print Type</th>" & _
                    "<th>Reprint Reason</th>" & _
                    "<th>Printed By</th>" & _
                    "<th>Quantity</th>" & _
                    "<th>Total Printed</th>" & _
                "</tr>"
    For i = 1 To rowCount
        emailBody = emailBody & "<tr>" & _
            "<td>" & filteredArray(i, 1) & "</td>" & _
            "<td>" & filteredArray(i, 14) & "</td>" & _
            "<td>" & filteredArray(i, 2) & "</td>" & _
            "<td>" & filteredArray(i, 5) & "</td>" & _
            "<td>" & filteredArray(i, 8) & "</td>" & _
            "<td>" & filteredArray(i, 18) & "</td>" & _
            "<td>" & filteredArray(i, 15) & "</td>" & _
            "<td>" & filteredArray(i, 16) & "</td>" & _
            "<td>" & filteredArray(i, 23) & "</td>" & _
            "<td>" & filteredArray(i, 6) & "</td>" & _
            "<td>" & filteredArray(i, 11) & "</td>" & _
            "</tr>"
    Next i
    emailBody = emailBody & "</table><p>Regards,<br>CPMU Label Printing Tool</p>"

    '--- Get or create Outlook instance ---
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    On Error GoTo 0
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If

    '--- Get the MAPI namespace and log on silently to the default profile ---
    Set ns = outlookApp.GetNamespace("MAPI")
    ns.Logon "", "", False, True   ' Silent logon using default profile

    '--- Prepare email subject and create the email ---
    mailSubject = "Label Print Summary"
    Set mailItem = outlookApp.CreateItem(0)
    With mailItem
        .To = recipientList
        .Subject = mailSubject
        .HTMLBody = emailBody

        ' Redirect the sent copy to Deleted Items instead of Sent Items.
        Set delFolder = ns.GetDefaultFolder(3) ' 3 = olFolderDeletedItems
        Set .SaveSentMessageFolder = delFolder

        .Send
    End With

    '--- Log the email in tblEmailHistory (optional) ---
    Set newRow = tblEmailHistory.ListRows.Add
    newRow.Range(1).value = recipientList
    newRow.Range(2).value = Now
    newRow.Range(3).value = "Printed Report"

    '--- Wait for 5 seconds to allow Outlook to move the sent email to Deleted Items ---
    Application.Wait Now + TimeValue("00:00:05")

    '--- Now, hunt down the email in Deleted Items and delete it ---
    nowTime = Now
    Set delFolder = ns.GetDefaultFolder(3)
    Set delItems = delFolder.Items

    filter = "[Subject] = '" & mailSubject & "'"
    Set filteredDelItems = delItems.Restrict(filter)
    For Each itm In filteredDelItems
        If DateDiff("n", itm.SentOn, Now) < 5 Then
            itm.Delete
        End If
    Next itm

Cleanup:
    Set mailItem = Nothing
    Set delFolder = Nothing
    Set delItems = Nothing
    Set filteredDelItems = Nothing
    Set foundItem = Nothing
    Set itm = Nothing
    Set ns = Nothing
    Set outlookApp = Nothing
    Set wsEmail = Nothing
    Set wsLabelPrints = Nothing
    Set tblEmailAdd = Nothing
    Set tblEmailHistory = Nothing
    Set tblLabelPrints = Nothing
    Erase filteredArray
    Set newRow = Nothing
    Set r = Nothing
End Sub

