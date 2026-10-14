Attribute VB_Name = "FileUtils"
Option Explicit

Public Function BuildFilePath(ByVal basePath As String, ParamArray parts() As Variant) As String
    Dim part As Variant
    Dim result As String
    
    result = basePath
    For Each part In parts
        result = result & "\" & CStr(part)
    Next
    
    BuildFilePath = result
End Function

Public Function GenerateTimestamp() As String
    GenerateTimestamp = Format(Now, "YYYYMMDD_HHMMSS")
End Function

Public Function GetWorkbookPath() As String
    GetWorkbookPath = ThisWorkbook.Path
End Function

Public Function BuildExportPath(ByVal fileName As String) As String
    Dim basePath As String
    basePath = GetWorkbookPath()
    
    BuildExportPath = BuildFilePath(basePath, "Exports", fileName)
End Function

Public Function IsFileLocked(ByVal filePath As String) As Boolean
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    Open filePath For Binary Access Write Lock Read Write As #ff
    If Err.Number <> 0 Then
        IsFileLocked = True
        Err.Clear
    Else
        IsFileLocked = False
        Close #ff
    End If
    On Error GoTo 0
End Function

Public Function OpenADOConnection(ByVal folderPath As String) As Object
    Dim conn As Object
    On Error GoTo ErrHandler

    ' --- Create ADO connection ---
    Set conn = CreateObject("ADODB.Connection")

    ' --- Attempt 32-bit connection ---
    On Error Resume Next
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & folderPath & ";" & _
              "Extended Properties=""text;HDR=Yes;FMT=Delimited;CharacterSet=65001"";"
    If Err.Number <> 0 Then
        ' --- If 32-bit connection fails, try 64-bit connection ---
        Err.Clear
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & folderPath & ";" & _
                  "Extended Properties=""text;HDR=Yes;FMT=Delimited;CharacterSet=65001"";"
    End If
    On Error GoTo ErrHandler

    ' --- Return the connection ---
    Set OpenADOConnection = conn
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "OpenADOConnection"
    Set OpenADOConnection = Nothing

CleanExit:
    Exit Function
End Function

Public Function HasBatchBeenPrinted(ByVal batchNumber As String) As Boolean
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object

    HasBatchBeenPrinted = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPath = CreateTemporaryFileCopy(FILE_FINISHED_GOODS_BATCH_AUDIT)
    
    If tempPath = "" Then
        GoTo CleanExit
    End If
    
    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)
    
    Set conn = OpenADOConnection(folderPath)

    If conn Is Nothing Then
        GoTo CleanExit
    End If

    sql = "SELECT COUNT(*) FROM [" & fileName & "] " & _
          "WHERE [" & FGBNM_HEADER_BATCH_NUMBER & "] = '" & Replace(batchNumber, "'", "''") & "' " & _
          "AND [" & FGBNM_HEADER_BATCH_STATUS & "] = '" & FGB_STATUS_PRINTED & "'"

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        HasBatchBeenPrinted = (rs.Fields(0).value > 0)
    End If

    GoTo CleanExit

ErrHandler:
    HasBatchBeenPrinted = False
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
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
    End If
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' DATA FILE PATH RESOLUTION
' ========================================================================
' Resolves a table file through PathResolver (repo DataFiles, customer
' DEV/DATA trees, or the workbook-local fallback) and validates it exists.
' ========================================================================

Public Function GetDataFilePath(fileName As String) As String
    On Error GoTo ErrHandler

    Dim fullPath As String

    fullPath = GetTableFilePath(fileName)

    ' --- Validate path exists ---
    If Not FileExists(fullPath) Then
        MsgBox "Data file not found: " & fullPath, vbCritical, "GetDataFilePath"
        GetDataFilePath = ""
        Exit Function
    End If

    GetDataFilePath = fullPath
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "GetDataFilePath"
    GetDataFilePath = ""
    GoTo CleanExit

CleanExit:
    Exit Function
End Function

' --- Helper: Resolve the drive letter from the current workbook path ---
Public Function GetWorkbookDrive() As String
    Dim wbPath As String
    Dim driveLetter As String
    
    wbPath = ThisWorkbook.Path
    
    If Len(wbPath) = 0 Then
        ' Workbook unsaved — default to common network drive
        GetWorkbookDrive = "G:"
        Exit Function
    End If
    
    ' Extract drive letter (first character + colon)
    If Len(wbPath) >= 2 Then
        driveLetter = Left$(wbPath, 2)
        If Right$(driveLetter, 1) = ":" Then
            GetWorkbookDrive = driveLetter
            Exit Function
        End If
    End If
    
    ' Fallback
    GetWorkbookDrive = "G:"
End Function

' --- Helper: Resolve a folder path based on workbook drive ---
Public Function GetFolderDrivePath(ByVal folderPath As String) As String
    Dim driveLetter As String
    driveLetter = GetWorkbookDrive()
    GetFolderDrivePath = driveLetter & folderPath
End Function

' ========================================================================
' Create a temporary copy of a source file for processing
' ========================================================================
Public Function CreateTemporaryFileCopy(ByVal sourceFileConstant As String, Optional ByVal sourceFolderPath As String = "") As String
    On Error GoTo ErrHandler
    
    Dim sourcePath As String
    Dim tempPath As String
    Dim tempFileName As String
    Dim copyAttempt As Integer
    Dim copySuccess As Boolean
    Dim fso As Object
    Dim tempFolder As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If sourceFolderPath = "" Then
        sourcePath = GetWorkbookTableFilePath(sourceFileConstant)
    Else
        sourcePath = sourceFolderPath
        If Right$(sourcePath, 1) <> "\" Then sourcePath = sourcePath & "\"
        sourcePath = sourcePath & sourceFileConstant
    End If
    
    If Not fso.FileExists(sourcePath) Then
        CreateTemporaryFileCopy = ""
        GoTo CleanExit
    End If
    
    tempFileName = GenerateUniqueTempFileName(sourceFileConstant)
    tempFolder = GetTempFolderPath()
    tempPath = tempFolder & tempFileName
    
    If Not EnsureFolderExists(tempFolder) Then
        CreateTemporaryFileCopy = ""
        GoTo CleanExit
    End If
    
    copySuccess = False
    For copyAttempt = 1 To 3
        On Error Resume Next
        fso.CopyFile sourcePath, tempPath, True
        If Err.Number = 0 Then
            copySuccess = True
            Exit For
        End If
        Err.Clear
        On Error GoTo ErrHandler
        Application.Wait Now + TimeValue("00:00:00") + 0.0000011574
    Next copyAttempt
    
    If copySuccess Then
        CreateTemporaryFileCopy = tempPath
    Else
        CreateTemporaryFileCopy = ""
    End If
    
    GoTo CleanExit

ErrHandler:
    CreateTemporaryFileCopy = ""
    GoTo CleanExit

CleanExit:
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' Generate unique temporary filename
' ========================================================================
Private Function GenerateUniqueTempFileName(ByVal originalFileName As String) As String
    Dim baseName As String
    Dim extension As String
    Dim timeStamp As String
    Dim uniqueId As String
    Dim dotPos As Long
    
    dotPos = InStrRev(originalFileName, ".")
    If dotPos > 0 Then
        baseName = Left$(originalFileName, dotPos - 1)
        extension = Mid$(originalFileName, dotPos)
    Else
        baseName = originalFileName
        extension = ""
    End If
    
    timeStamp = Format$(Now, "yyyymmdd_hhnnss_") & Format$(Timer * 1000 Mod 1000, "000")
    uniqueId = CStr(Int(Rnd() * 10000))
    
    GenerateUniqueTempFileName = baseName & "_" & timeStamp & "_" & uniqueId & extension
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    
    On Error GoTo ErrHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    
    GoTo CleanExit
    
ErrHandler:
    FileExists = False
    
CleanExit:
    If Not fso Is Nothing Then Set fso = Nothing
End Function

Public Function EnsureFolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    
    On Error GoTo ErrHandler
    
    ' --- Check if folder already exists ---
    If FolderExists(folderPath) Then
        EnsureFolderExists = True
        Exit Function
    End If
    
    ' --- Create folder if it doesn't exist ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder folderPath
    EnsureFolderExists = True
    
    GoTo CleanExit
    
ErrHandler:
    EnsureFolderExists = False
    
CleanExit:
    If Not fso Is Nothing Then Set fso = Nothing
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    
    On Error GoTo ErrHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
    
    GoTo CleanExit
    
ErrHandler:
    FolderExists = False
    
CleanExit:
    If Not fso Is Nothing Then Set fso = Nothing
End Function

' Environment detection (GetEnvironmentFromPath) lives in PathResolver.

' ========================================================================
' GET UNSCANNED BATCHES
' ========================================================================
' Queries the batch master file via ADO to retrieve batch numbers where
' Scanned field is not "x". Returns array of batch numbers.
' ========================================================================

Public Function GetUnscannedBatches() As String()
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim rs As Object
    Dim tempPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim sql As String
    Dim fso As Object
    Dim batches() As String
    Dim tempBatches() As String
    Dim i As Long
    Dim batchCount As Long

    ReDim batches(-1 To -1)

    Set fso = CreateObject("Scripting.FileSystemObject")

    tempPath = CreateTemporaryFileCopy(FILE_FINISHED_GOODS_BATCH_MASTER)

    If tempPath = "" Then
        GoTo CleanExit
    End If

    folderPath = fso.GetParentFolderName(tempPath)
    fileName = fso.GetFileName(tempPath)

    Set conn = OpenADOConnection(folderPath)

    If conn Is Nothing Then
        GoTo CleanExit
    End If

    sql = "SELECT [" & FGBNM_HEADER_BATCH_NUMBER & "] " & _
          "FROM [" & fileName & "] " & _
          "WHERE ([" & FGBNM_HEADER_SCANNED & "] IS NULL OR [" & FGBNM_HEADER_SCANNED & "] <> 'x')"

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        ReDim tempBatches(0 To 999)
        batchCount = 0
        
        Do While Not rs.EOF
            tempBatches(batchCount) = Trim$(CStr(rs.Fields(0).value))
            batchCount = batchCount + 1
            
            If batchCount > UBound(tempBatches) Then
                ReDim Preserve tempBatches(0 To UBound(tempBatches) + 1000)
            End If
            
            rs.MoveNext
        Loop
        
        If batchCount > 0 Then
            ReDim batches(0 To batchCount - 1)
            For i = 0 To batchCount - 1
                batches(i) = tempBatches(i)
            Next i
        End If
    End If

    GetUnscannedBatches = batches
    GoTo CleanExit

ErrHandler:
    ReDim batches(-1 To -1)
    GetUnscannedBatches = batches
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
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    End If
    Set fso = Nothing
    Exit Function
End Function

' ========================================================================
' CHECK BATCH PDFS EXIST
' ========================================================================
' Checks if PDF files exist in the completed archive folder for given
' batch numbers. Returns array of batch numbers with confirmed files.
' ========================================================================

Public Function CheckBatchPDFsExist(ByRef batchNumbers() As String) As String()
    On Error GoTo ErrHandler

    Dim fso As Object
    Dim archivePath As String
    Dim pdfPath As String
    Dim foundBatches() As String
    Dim foundCount As Long
    Dim i As Long
    Dim inputLBound As Long
    Dim inputUBound As Long
    Dim firstFound As Boolean

    On Error Resume Next
    inputLBound = LBound(batchNumbers)
    inputUBound = UBound(batchNumbers)
    If Err.Number <> 0 Then
        ReDim foundBatches(-1 To -1)
        CheckBatchPDFsExist = foundBatches
        Exit Function
    End If
    On Error GoTo ErrHandler

    foundCount = 0
    firstFound = False

    Set fso = CreateObject("Scripting.FileSystemObject")

    archivePath = GetWorkbookDrive() & PRODUCTION_ARCHIVE_COMPLETED_PMI_PATH

    For i = inputLBound To inputUBound
        pdfPath = archivePath & batchNumbers(i) & ".pdf"

        If fso.FileExists(pdfPath) Then
            If Not firstFound Then
                ReDim foundBatches(0 To 0)
                firstFound = True
            Else
                ReDim Preserve foundBatches(0 To foundCount)
            End If
            
            foundBatches(foundCount) = batchNumbers(i)
            foundCount = foundCount + 1
        End If
    Next i

    If Not firstFound Then
        ReDim foundBatches(-1 To -1)
    End If

    CheckBatchPDFsExist = foundBatches
    GoTo CleanExit

ErrHandler:
    ReDim foundBatches(-1 To -1)
    CheckBatchPDFsExist = foundBatches
    GoTo CleanExit

CleanExit:
    Set fso = Nothing
    Exit Function
End Function