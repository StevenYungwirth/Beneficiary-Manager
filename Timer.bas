Attribute VB_Name = "Timer"
Option Explicit
Private LogStarted As Boolean
Private TimerRunning As Boolean
Private StartingTime As Double
Private TotalElapsedTime As Double

Public Sub TimeStart()
    'Check if the timer's already running; don't change anything if it is
    If Not TimerRunning Then
        StartingTime = Now
        TimerRunning = True
        Debug.Print "Started at: " & StartingTime
    End If
End Sub

Public Sub TimeEnd()
    'Get the current time and compare it to when the timer started
    Dim ElapsedTime As Double
    ElapsedTime = Now - StartingTime
    TotalElapsedTime = TotalElapsedTime + ElapsedTime
    Debug.Print "Ended at: " & Now
    Debug.Print "Elapsed time: " & ElapsedTime
    Debug.Print "Total Elapsed time: " & Round(TotalElapsedTime * 60 * 60 * 24, 2)
    
    'Stop the timer
    TimerRunning = False
End Sub
