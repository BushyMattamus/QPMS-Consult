Attribute VB_Name = "AppStateUtils"
Option Explicit

Public Sub DisableAppUpdates()
    ' Disable updates
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

Public Sub RestoreAppState()
    ' Restore to original state
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Public Function ExecuteWithDisabledUpdates(ByVal procedure As String) As Boolean
    On Error GoTo ErrorHandler
    
    DisableAppUpdates
    Application.Run procedure
    ExecuteWithDisabledUpdates = True
    
ExitFunction:
    RestoreAppState
    Exit Function
    
ErrorHandler:
    ExecuteWithDisabledUpdates = False
    Resume ExitFunction
End Function