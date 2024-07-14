Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateTimer()
    Dim t As TimerObject: Set t = New TimerObject
    Set CreateTimer = t
End Function

Public Function CreateBooster()
    Dim b As PerformanceBooster: Set b = New PerformanceBooster
    Set CreateBooster = b
End Function
