Attribute VB_Name = "PerfTimer"
Option Explicit

' ========================================================================
' PERFORMANCE TIMER
' ========================================================================
' Shared high-resolution timing for Immediate window diagnostics.
' ========================================================================

' Currency is used for the 64-bit counter so one signature works on both
' 32-bit and 64-bit VBA. The raw integer is read back as CDbl(value) * 10000.
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Private m_frequency As Double

Private Function PerfTimerFrequency() As Double
    Dim f As Currency
    
    If m_frequency = 0 Then
        QueryPerformanceFrequency f
        m_frequency = CDbl(f) * 10000#
    End If
    
    PerfTimerFrequency = m_frequency
End Function

Public Function PerfTimerNow() As Double
    Dim c As Currency
    
    QueryPerformanceCounter c
    PerfTimerNow = CDbl(c) * 10000# / PerfTimerFrequency()
End Function

Public Sub PerfTimerLog(ByVal tag As String, ByVal label As String, ByVal startSec As Double)
    Debug.Print "[" & tag & " TIMER] " & label & ": " & _
                Format$((PerfTimerNow() - startSec) * 1000#, "0.000") & " ms"
End Sub
