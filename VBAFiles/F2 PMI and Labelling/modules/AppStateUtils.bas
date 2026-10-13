Attribute VB_Name = "AppStateUtils"
Option Explicit

Public Sub DisableAppUpdates()
    ' Disable updates
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
End Sub

Public Sub RestoreAppState()
    ' Restore to original state
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
End Sub