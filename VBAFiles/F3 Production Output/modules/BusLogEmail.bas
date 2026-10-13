Attribute VB_Name = "BusLogEmail"
Option Explicit
''' Joins only the non-blank items in a ParamArray, separated by ", "
Private Function JoinNonEmpty(ParamArray items() As Variant) As String
    Dim arr() As String
    Dim i As Long, cnt As Long
    cnt = 0
    
    For i = LBound(items) To UBound(items)
        If Len(Trim(items(i) & "")) > 0 Then
            ReDim Preserve arr(0 To cnt)
            arr(cnt) = items(i)
            cnt = cnt + 1
        End If
    Next i
    
    If cnt > 0 Then
        JoinNonEmpty = Join(arr, ", ")
    Else
        JoinNonEmpty = ""
    End If
End Function


Sub GenerateSummaryEmail_Silent_DeleteSent()
    ' Worksheet variables
    Dim wsEmail As Worksheet
    Dim wstblProdOutput As Worksheet

    ' ListObject variables
    Dim tblEmailAdd As ListObject
    Dim tblEmailHistory As ListObject
    Dim tblProdOutput As ListObject

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
    Set wstblProdOutput = ThisWorkbook.Sheets("tblProdOutput")

    Set tblEmailAdd = wsEmail.ListObjects("tblEmailAdd")
    Set tblProdOutput = wstblProdOutput.ListObjects("tblProdOutput")
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
        If r.Range(3).value = "Production Output" Then
            If r.Range(2).value > lastEmailDate Then
                lastEmailDate = r.Range(2).value
            End If
        End If
    Next r

    '--- Extract data from tblProdOutput ---
    If tblProdOutput.DataBodyRange Is Nothing Then GoTo Cleanup
    lastRow = tblProdOutput.DataBodyRange.Rows.Count
    If lastRow = 0 Then GoTo Cleanup

    dataArray = tblProdOutput.DataBodyRange.value

    '--- Filter by Record Saved date time ---
    ReDim filteredArray(1 To UBound(dataArray, 1), 1 To UBound(dataArray, 2))
    rowCount = 0

    For i = 1 To UBound(dataArray, 1)
        If Not IsError(dataArray(i, 35)) Then ' column 35 contains the data saving date
            If IsDate(dataArray(i, 35)) Then
                Dim recD As Date
                recD = CDate(dataArray(i, 35))
                If recD > lastEmailDate Then
                    rowCount = rowCount + 1
                    For j = 1 To UBound(dataArray, 2)
                        filteredArray(rowCount, j) = dataArray(i, j)
                    Next j
                End If
            End If
        End If
    Next i
    If rowCount = 0 Then GoTo Cleanup

    '--- Build the HTML table for the email body ---
    emailBody = "<p>Dear Team,</p>" & _
            "<p>Please find below the Production Output report from " & Format(lastEmailDate, "d/MM/yyyy") & " at " & Format(lastEmailDate, "h:mm:ss AM/PM") & ".</p>" & _
            "<table border='1' cellspacing='0' cellpadding='5'>" & _
            "<tr>" & _
                "<th>Actual Start Date</th>" & _
                "<th>Planned Start Date</th>" & _
                "<th>Product Description</th>" & _
                "<th>Quick Code</th>" & _
                "<th>CP PMI Code</th>" & _
                "<th>Start Room</th>" & _
                "<th>Finish Room</th>" & _
                "<th>Batch</th>" & _
                "<th>Output Qty (ea)</th>" & _
                "<th>Notes</th>" & _
            "</tr>"
    For i = 1 To rowCount
        emailBody = emailBody & "<tr>" & _
            "<td>" & filteredArray(i, 1) & "</td>" & _
            "<td>" & filteredArray(i, 2) & "</td>" & _
            "<td>" & filteredArray(i, 3) & "</td>" & _
            "<td>" & filteredArray(i, 4) & "</td>" & _
            "<td>" & filteredArray(i, 5) & "</td>" & _
            "<td>" & filteredArray(i, 8) & "</td>" & _
            "<td>" & filteredArray(i, 9) & "</td>" & _
            "<td>" & filteredArray(i, 10) & "</td>" & _
            "<td>" & filteredArray(i, 11) & "</td>" & _
            "<td>" & filteredArray(i, 33) & "</td>" & _
            "</tr>"
    Next i
    emailBody = emailBody & "</table><p>Regards,<br>Production Output Summary</p>"

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
    mailSubject = "Production Output Summary"
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
    
    wsEmail.Unprotect

    '--- Log the email in tblEmailHistory (optional) ---
    Set newRow = tblEmailHistory.ListRows.Add
    newRow.Range(1).value = recipientList
    newRow.Range(2).value = Now
    newRow.Range(3).value = "Production Output"
    
    wsEmail.Protect

    '--- Wait for 5 seconds to allow Outlook to move the sent email to Deleted Items ---
    Application.Wait Now + TimeValue("00:00:05")

    '--- Now, hunt down the email in Deleted Items and delete it ---
    nowTime = Now
    Set delFolder = ns.GetDefaultFolder(3)
    Set delItems = delFolder.items

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
    Set wstblProdOutput = Nothing
    Set tblEmailAdd = Nothing
    Set tblEmailHistory = Nothing
    Set tblProdOutput = Nothing
    Erase filteredArray
    Set newRow = Nothing
    Set r = Nothing
End Sub
