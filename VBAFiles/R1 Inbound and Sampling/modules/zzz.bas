Attribute VB_Name = "zzz"
'=====================================================
' Easter Egg – CPMU Secret Surprise
'=====================================================
' Trigger combo: Ctrl + Shift + E
' To enable: Place this in a standard module.
'=====================================================

Option Explicit

Sub ActivateEasterEgg()
    Dim words As Variant
    Dim i As Integer
    
    words = Array("why", "did", "you", "do", "this", "to", "me")

    Application.StatusBar = "Activating secret mode..."
    Pause 0.5

    For i = LBound(words) To UBound(words)
        Application.StatusBar = Application.StatusBar & " " & words(i)
        DoEvents
        Pause 0.5
    Next i

    Application.StatusBar = False

    MsgBox "Congratulations, explorer!" & vbCrLf & _
           "You’ve uncovered the hidden Easter Egg..." & vbCrLf & vbCrLf & _
           "(¬?¬) Mmm… donuts." & vbCrLf & vbCrLf & _
           "Now go get yourself a coffee — you’ve earned it!", _
           vbInformation + vbOKOnly, "Secret Found"
End Sub

Private Sub Pause(seconds As Single)
    Dim EndTime As Single
    EndTime = Timer + seconds
    Do While Timer < EndTime
        DoEvents
    Loop
End Sub

