Option Explicit

' Module1

Sub PrintGCD
    Dim a As Long: a = 3120
    Dim b As Long: b = 45
    Dim result As Long
    result = myGCD(a, b)
    Debug.Print("GCD=" & result)
End Sub

Function myGCD(ByVal a As Long, ByVal b As Long)
    Dim t As Long
    While b > 0
        t = a Mod b
        a = b
        b = t
    Wend
    myGCD = a
End Function