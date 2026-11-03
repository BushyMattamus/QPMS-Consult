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
              "Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
    If Err.Number <> 0 Then
        ' --- If 32-bit connection fails, try 64-bit connection ---
        Err.Clear
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & folderPath & ";" & _
                  "Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
    End If
    On Error GoTo ErrHandler

    ' --- Return the connection ---
    Set OpenADOConnection = conn
    Exit Function

ErrHandler:
    ' --- Handle errors ---
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "OpenADOConnection"
    Set OpenADOConnection = Nothing
End Function

' ========================================================================
' Create a temporary copy of a source file for processing
' ========================================================================
Public Function CreateTemporaryFileCopy(ByVal sourceFileConstant As String, Optional ByVal sourceFolderPath As String = "") As String
    On Error GoTo ErrHandler
    
    Dim sourcePath As String
    Dim tempPath As String
    Dim tempFileName As String
    Dim tempFolder As String
    Dim copyAttempt As Integer
    Dim copySuccess As Boolean
    Dim fso As Object
    
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
    Dim timestamp As String
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
    
    timestamp = Format$(Now, "yyyymmdd_hhnnss_") & Format$(Timer * 1000 Mod 1000, "000")
    uniqueId = CStr(Int(Rnd() * 10000))
    
    GenerateUniqueTempFileName = baseName & "_" & timestamp & "_" & uniqueId & extension
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
