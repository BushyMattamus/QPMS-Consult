Attribute VB_Name = "FormUtils"
Option Explicit

'--- Unprotects the active worksheet (no password required)
Public Sub UnprotectActiveSheet()
    On Error Resume Next
    ActiveSheet.Unprotect
    On Error GoTo 0
End Sub

'--- Protects the active worksheet (no password required)
Public Sub ProtectActiveSheet()
    On Error Resume Next
    If PROD_DEV <> "DEV" Then
        ActiveSheet.Protect _
            AllowFiltering:=False, _
            AllowSorting:=False
    End If
    On Error GoTo 0
End Sub

'--- Locks workbook structure
Public Sub LockWorkbook()
    On Error Resume Next
    ActiveWorkbook.Protect Structure:=True, Windows:=True
    On Error GoTo 0
End Sub

' --- Unlock the workbook layout ---
Public Sub UnlockWorkbook()
    On Error Resume Next
    ActiveWorkbook.Protect Structure:=False, Windows:=False
    On Error GoTo 0
End Sub

Public Function GetFieldValByLabelC(ByVal labelText As String) As String
    ' Finds the label in column C (any row) and returns the value from column D, same row.
    ' Match is case-insensitive and ignores surrounding spaces.
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    Dim r As Long
    For r = 1 To lastRow
        If LCase$(Trim$(ws.Cells(r, "C").value)) = LCase$(Trim$(labelText)) Then
            GetFieldValByLabelC = Trim$(ws.Cells(r, "D").value)
            Exit Function
        End If
    Next r
    GetFieldValByLabelC = ""
End Function

Public Function YesNoToYN(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    If t = "yes" Then
        YesNoToYN = "Y"
    ElseIf t = "no" Then
        YesNoToYN = "N"
    Else
        ' If blank or unrecognized, default to "N" (adjust if you prefer blank)
        YesNoToYN = "N"
    End If
End Function

' --- Append QC Number record to audit trail file ---
Public Sub AppendToQCNumberAudit(ByVal recordValues As Variant)
    Dim auditPath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    
    On Error GoTo ErrorHandler
    
    ' Define audit file path
    auditPath = GetWorkbookTableFilePath(FILE_QC_NUMBER_CHANGES)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Format record for CSV
    If IsArray(recordValues) Then
        line = Join(recordValues, ",")
    Else
        line = recordValues ' In case a pre-formatted string is passed
    End If
    
    ' Append record to audit file
    Set ts = fso.OpenTextFile(auditPath, 8, True) ' 8 = ForAppending
    ts.WriteLine line
    ts.Close
    
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error writing to audit trail: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
End Sub

'===== Fuzzy Search Functions =====
Public Function FuzzyMatch(ByVal searchTerm As String, ByVal targetText As String) As Double
    ' Returns a score between 0 and 1 (1 = perfect match)
    If Len(searchTerm) = 0 Or Len(targetText) = 0 Then
        FuzzyMatch = 0
        Exit Function
    End If
    
    searchTerm = UCase$(Trim$(searchTerm))
    targetText = UCase$(Trim$(targetText))
    
    ' Exact match
    If searchTerm = targetText Then
        FuzzyMatch = 1
        Exit Function
    End If
    
    ' Contains check (high priority for user experience)
    If InStr(targetText, searchTerm) > 0 Then
        ' Score based on how much of the target the search term represents
        FuzzyMatch = 0.8 + (Len(searchTerm) / Len(targetText)) * 0.2
        Exit Function
    End If
    
    ' Check if search term starts any word in target
    Dim words() As String
    words = Split(targetText, " ")
    Dim i As Long
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 And InStr(words(i), searchTerm) = 1 Then
            FuzzyMatch = 0.7
            Exit Function
        End If
    Next i
    
    ' Check if any word starts with search term
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 And InStr(words(i), searchTerm) > 0 Then
            FuzzyMatch = 0.6
            Exit Function
        End If
    Next i
    
    ' Use Levenshtein distance for similarity scoring
    Dim maxLen As Long
    maxLen = IIf(Len(searchTerm) > Len(targetText), Len(searchTerm), Len(targetText))
    
    If maxLen > 0 Then
        Dim distance As Long
        distance = LevenshteinDistance(searchTerm, targetText)
        
        ' Convert distance to similarity score (0 to 1)
        Dim similarity As Double
        similarity = 1 - (distance / maxLen)
        
        ' Only return if similarity is reasonable
        If similarity >= 0.3 Then
            FuzzyMatch = similarity * 0.5 ' Scale down Levenshtein matches
        Else
            FuzzyMatch = 0
        End If
    Else
        FuzzyMatch = 0
    End If
End Function

Function LevenshteinDistance(s1 As String, s2 As String) As Long
    Dim d() As Long
    Dim i As Long, j As Long, cost As Long
    Dim len1 As Long, len2 As Long

    len1 = Len(s1)
    len2 = Len(s2)

    ReDim d(len1, len2)

    For i = 0 To len1
        d(i, 0) = i
    Next i

    For j = 0 To len2
        d(0, j) = j
    Next j

    For i = 1 To len1
        For j = 1 To len2
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            d(i, j) = WorksheetFunction.Min(d(i - 1, j) + 1, _
                                             d(i, j - 1) + 1, _
                                             d(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = d(len1, len2)
End Function

' Alternative function for partial word matching within longer strings
Public Function FuzzyMatchPartial(ByVal searchTerm As String, ByVal targetText As String) As Double
    ' This version checks if search term closely matches any part of target text
    If Len(searchTerm) = 0 Or Len(targetText) = 0 Then
        FuzzyMatchPartial = 0
        Exit Function
    End If
    
    searchTerm = UCase$(Trim$(searchTerm))
    targetText = UCase$(Trim$(targetText))
    
    ' Direct contains check first
    If InStr(targetText, searchTerm) > 0 Then
        FuzzyMatchPartial = 0.9
        Exit Function
    End If
    
    ' Check Levenshtein distance against each word
    Dim words() As String
    words = Split(targetText, " ")
    Dim i As Long
    Dim bestScore As Double
    bestScore = 0
    
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            Dim maxLen As Long
            maxLen = IIf(Len(searchTerm) > Len(words(i)), Len(searchTerm), Len(words(i)))
            
            If maxLen > 0 Then
                Dim distance As Long
                distance = LevenshteinDistance(searchTerm, words(i))
                Dim similarity As Double
                similarity = 1 - (distance / maxLen)
                
                If similarity > bestScore Then
                    bestScore = similarity
                End If
            End If
        End If
    Next i
    
    FuzzyMatchPartial = bestScore
End Function

Public Sub UpdateFuzzySearchResults(ByVal searchCell As Range, ByVal outputStartCell As Range, Optional ByVal maxResults As Long = 20)
    Dim searchTerm As String
    Dim wsData As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim results As Collection
    Dim i As Long
    Dim rmCode As String
    Dim stdName As String
    Dim altName As String
    Dim score1 As Double, score2 As Double, maxScore As Double
    
    On Error GoTo ErrorHandler
    
    searchTerm = Trim$(searchCell.value)
    Set wsOutput = searchCell.Worksheet
    Set wsData = ThisWorkbook.Worksheets("RM_Master")
    
    ' Clear previous results (next 30 rows, now 4 columns: RM Code, Std Name, Alt Name, Score)
    wsOutput.Range(outputStartCell, outputStartCell.Offset(29, 3)).ClearContents
    
    ' Unlock the output range so it remains accessible when worksheet is protected
    wsOutput.Range(outputStartCell, outputStartCell.Offset(29, 3)).Locked = False
    
    If Len(searchTerm) < 2 Then Exit Sub ' Don't search until at least 2 characters
    
    Set results = New Collection
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row
    
    ' Get column positions
    Dim colRM As Long, colStd As Long, colAlt As Long
    colRM = GetHeaderCol(wsData, "RM_Code", 1)
    colStd = GetHeaderCol(wsData, "Std_Name", 1)
    colAlt = GetHeaderCol(wsData, "Alt_Name", 1)
    
    If colRM = 0 Or colStd = 0 Then Exit Sub
    
    ' Search through data
    For i = 2 To lastRow ' Skip header row
        rmCode = Trim$(wsData.Cells(i, colRM).value)
        stdName = Trim$(wsData.Cells(i, colStd).value)
        altName = IIf(colAlt > 0, Trim$(wsData.Cells(i, colAlt).value), "")
        
        If Len(rmCode) > 0 And Len(stdName) > 0 Then
            ' Calculate fuzzy match scores using both methods
            score1 = FuzzyMatch(searchTerm, stdName)
            score2 = IIf(Len(altName) > 0, FuzzyMatch(searchTerm, altName), 0)
            
            ' Also try partial matching for better results
            Dim partialScore1 As Double, partialScore2 As Double
            partialScore1 = FuzzyMatchPartial(searchTerm, stdName)
            partialScore2 = IIf(Len(altName) > 0, FuzzyMatchPartial(searchTerm, altName), 0)
            
            ' Use the best score from either method
            score1 = IIf(score1 > partialScore1, score1, partialScore1)
            score2 = IIf(score2 > partialScore2, score2, partialScore2)
            
            maxScore = IIf(score1 > score2, score1, score2)
            
            ' Lower threshold for better matches
            If maxScore > 0.3 Then
                ' Store as array: [score, rmCode, description]
                Dim resultItem(3) As Variant
                resultItem(0) = maxScore
                resultItem(1) = rmCode
                resultItem(2) = stdName
                resultItem(3) = altName
                
                results.Add resultItem
            End If
        End If
    Next i
    
    ' Convert Collection to Array for easier sorting
    If results.Count > 0 Then
        Dim resultArray() As Variant
        ReDim resultArray(1 To results.Count, 1 To 4)
        
        For i = 1 To results.Count
            resultArray(i, 1) = results(i)(0) ' Score
            resultArray(i, 2) = results(i)(1) ' RM Code
            resultArray(i, 3) = results(i)(2) ' Standard Name
            resultArray(i, 4) = results(i)(3) ' Alternate Name
        Next i
        
        ' Simple bubble sort on the array
        Dim j As Long
        Dim temp1 As Variant, temp2 As Variant, temp3 As Variant, temp4 As Variant
        For i = 1 To results.Count - 1
            For j = i + 1 To results.Count
                If resultArray(i, 1) < resultArray(j, 1) Then ' Sort by score descending
                    ' Swap rows
                    temp1 = resultArray(i, 1): temp2 = resultArray(i, 2): temp3 = resultArray(i, 3): temp4 = resultArray(i, 4)
                    resultArray(i, 1) = resultArray(j, 1): resultArray(i, 2) = resultArray(j, 2): resultArray(i, 3) = resultArray(j, 3): resultArray(i, 4) = resultArray(j, 4)
                    resultArray(j, 1) = temp1: resultArray(j, 2) = temp2: resultArray(j, 3) = temp3: resultArray(j, 4) = temp4
                End If
            Next j
        Next i
        
        ' Output top results
        Dim outputRow As Long
        outputRow = 0
        Dim maxOutput As Long
        maxOutput = IIf(results.Count > maxResults, maxResults, results.Count)
        
        ' Format the RM Code output column as Text before writing values
        wsOutput.Range(wsOutput.Cells(outputStartCell.row, outputStartCell.Column), _
                      wsOutput.Cells(outputStartCell.row + maxOutput - 1, outputStartCell.Column)).NumberFormat = "@"
        
        For i = 1 To maxOutput
            wsOutput.Cells(outputStartCell.row + outputRow, outputStartCell.Column).value = RM_Pad4(resultArray(i, 2)) ' RM Code with Pad4 formatting
            wsOutput.Cells(outputStartCell.row + outputRow, outputStartCell.Column + 1).value = resultArray(i, 3) ' Standard Name
            wsOutput.Cells(outputStartCell.row + outputRow, outputStartCell.Column + 2).value = resultArray(i, 4) ' Alternate Name
            wsOutput.Cells(outputStartCell.row + outputRow, outputStartCell.Column + 3).value = Format(resultArray(i, 1), "0.00") ' Score
            outputRow = outputRow + 1
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    ' For debugging, you might want to temporarily show errors
    MsgBox "Search Error: " & Err.Description
End Sub

'===== Helper function to get column by header =====
Public Function GetHeaderCol(ByVal ws As Worksheet, ByVal headerName As String, ByVal headerRow As Long) As Long
    Dim lastCol As Long
    Dim i As Long
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If Trim$(UCase$(ws.Cells(headerRow, i).value)) = Trim$(UCase$(headerName)) Then
            GetHeaderCol = i
            Exit Function
        End If
    Next i
    GetHeaderCol = 0
End Function

Public Function ValidateNoCommas(ByVal value As String, ByVal fieldName As String) As Boolean
    If InStr(1, value, ",") > 0 Then
        MsgBox "Error: " & fieldName & " cannot contain commas. Please remove any commas and try again.", vbCritical
        ValidateNoCommas = False
    Else
        ValidateNoCommas = True
    End If
End Function
