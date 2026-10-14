Attribute VB_Name = "BackupUtils"
Option Explicit

Public Sub BackupFrontend()
    Dim fso As Object
    Dim backupFolder As String
    Dim backupPath As String
    Dim todayBackup As String
    Dim sourceFile As String
    Dim currentDay As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    sourceFile = ThisWorkbook.FullName
    
    backupFolder = ThisWorkbook.Path & "\Backups\Frontend"
    
    ' --- Ensure backup folder structure exists ---
    EnsureFolderExists ThisWorkbook.Path & "\Backups"
    EnsureFolderExists backupFolder
    
    todayBackup = Format(Date, "yyyy-mm-dd")
    backupPath = backupFolder & "\" & todayBackup & " - " & fso.GetBaseName(sourceFile) & ".xlsm"
    
    On Error Resume Next
    fso.CopyFile sourceFile, backupPath, True
    On Error GoTo 0

    currentDay = Day(Date)
    If currentDay = 5 Or currentDay = 15 Or currentDay = 25 Then
        CleanupOldBackups backupFolder, 30
    End If
    
    Set fso = Nothing
End Sub

Public Sub BackupDataTables()
    Dim fso As Object
    Dim tablesPath As String
    Dim backupFolder As String
    Dim previousBackupsFolder As String
    Dim dataBackupFolder As String
    Dim file As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim currentBackupFolder As String
    Dim lastBackupTime As Date
    Dim folderCount As Integer
    Dim currentDay As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tablesPath = GetWorkbookTablesPath()
    backupFolder = ThisWorkbook.Path & "\Backups\Data"
    previousBackupsFolder = backupFolder & "\Previous Backups"
    
    ' --- Ensure backup folder structure exists ---
    EnsureFolderExists ThisWorkbook.Path & "\Backups"
    EnsureFolderExists backupFolder
    EnsureFolderExists previousBackupsFolder
    
    Set folder = fso.GetFolder(backupFolder)
    folderCount = 0
    currentBackupFolder = ""
    
    For Each subFolder In folder.SubFolders
        If subFolder.Name <> "Previous Backups" Then
            folderCount = folderCount + 1
            currentBackupFolder = subFolder.Path
        End If
    Next subFolder
    
    If folderCount = 1 And currentBackupFolder <> "" Then
        lastBackupTime = GetFolderDateTimeFromName(fso.GetFileName(currentBackupFolder))
        
        If DateDiff("n", lastBackupTime, Now) < 60 Then
            Set fso = Nothing
            Exit Sub
        End If
        
        On Error Resume Next
        fso.MoveFolder currentBackupFolder, previousBackupsFolder & "\"
        On Error GoTo 0
    End If
    
    dataBackupFolder = backupFolder & "\" & Format(Now, "yyyy-mm-dd HH-mm-ss")
    
    On Error Resume Next
    fso.CreateFolder (dataBackupFolder)
    On Error GoTo 0
    
    If FolderExists(tablesPath) Then
        On Error Resume Next
        For Each file In fso.GetFolder(tablesPath).Files
            If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
                fso.CopyFile file.Path, dataBackupFolder & "\" & file.Name, True
            End If
        Next file
        On Error GoTo 0
    End If
    
    currentDay = Day(Date)
    If currentDay = 5 Or currentDay = 15 Or currentDay = 25 Then
        CleanupOldPreviousBackups previousBackupsFolder, 30
    End If
    
    Set folder = Nothing
    Set fso = Nothing
End Sub

Private Function GetFolderDateTimeFromName(ByVal folderName As String) As Date
    Dim yearPart As String
    Dim monthPart As String
    Dim dayPart As String
    Dim hourPart As String
    Dim minPart As String
    Dim secPart As String
    
    On Error Resume Next
    
    If Len(folderName) >= 19 Then
        yearPart = Mid(folderName, 1, 4)
        monthPart = Mid(folderName, 6, 2)
        dayPart = Mid(folderName, 9, 2)
        hourPart = Mid(folderName, 12, 2)
        minPart = Mid(folderName, 15, 2)
        secPart = Mid(folderName, 18, 2)
        
        If IsNumeric(yearPart) And IsNumeric(monthPart) And IsNumeric(dayPart) And _
           IsNumeric(hourPart) And IsNumeric(minPart) And IsNumeric(secPart) Then
            GetFolderDateTimeFromName = DateSerial(CInt(yearPart), CInt(monthPart), CInt(dayPart)) + _
                                        TimeSerial(CInt(hourPart), CInt(minPart), CInt(secPart))
        Else
            GetFolderDateTimeFromName = CDate("1900-01-01 00:00:00")
        End If
    Else
        GetFolderDateTimeFromName = CDate("1900-01-01 00:00:00")
    End If
    
    On Error GoTo 0
End Function

Private Sub CleanupOldPreviousBackups(ByVal previousBackupsFolder As String, ByVal retentionDays As Integer)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim cutoffDate As Date
    Dim folderDate As Date
    Dim folderName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not FolderExists(previousBackupsFolder) Then
        Set fso = Nothing
        Exit Sub
    End If
    
    cutoffDate = Date - retentionDays
    Set folder = fso.GetFolder(previousBackupsFolder)
    
    On Error Resume Next
    For Each subFolder In folder.SubFolders
        folderName = subFolder.Name
        If Len(folderName) >= 10 Then
            folderDate = DateSerial(CInt(Mid(folderName, 1, 4)), _
                                    CInt(Mid(folderName, 6, 2)), _
                                    CInt(Mid(folderName, 9, 2)))
            If folderDate < cutoffDate Then
                fso.DeleteFolder subFolder.Path, True
            End If
        End If
    Next subFolder
    On Error GoTo 0
    
    Set folder = Nothing
    Set fso = Nothing
End Sub

Private Function IsNumeric(ByVal value As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsEmpty(value + 0)
    On Error GoTo 0
End Function

Private Sub CleanupOldBackups(ByVal backupFolder As String, ByVal retentionDays As Integer)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim cutoffDate As Date
    Dim fileDate As Date
    Dim fileName As String
    Dim yearPart As String
    Dim monthPart As String
    Dim dayPart As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not FolderExists(backupFolder) Then
        Set fso = Nothing
        Exit Sub
    End If
    
    cutoffDate = Date - retentionDays
    Set folder = fso.GetFolder(backupFolder)
    
    On Error Resume Next
    For Each file In folder.Files
        fileName = file.Name
        
        If LCase(fso.GetExtensionName(fileName)) = "xlsm" And Len(fileName) >= 10 Then
            yearPart = Mid(fileName, 1, 4)
            monthPart = Mid(fileName, 6, 2)
            dayPart = Mid(fileName, 9, 2)
            
            If IsNumeric(yearPart) And IsNumeric(monthPart) And IsNumeric(dayPart) Then
                fileDate = DateSerial(CInt(yearPart), CInt(monthPart), CInt(dayPart))
                
                If fileDate < cutoffDate Then
                    fso.DeleteFile file.Path, True
                End If
            End If
        End If
    Next file
    On Error GoTo 0
    
    Set folder = Nothing
    Set fso = Nothing
End Sub


