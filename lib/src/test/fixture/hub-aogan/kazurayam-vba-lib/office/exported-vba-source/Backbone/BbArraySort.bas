Attribute VB_Name = "BbArraySort"
Option Explicit

' 配列をソートする

' https://www.tipsfound.com/vba/02020


Public Sub InsertionSort(ByRef data As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Variant
    Dim k As Variant
    Dim t As Variant
    
    For i = low + 1 To high
        t = data(i)
        If data(i - 1) > t Then
            k = i
            Do While k > low
                If data(k - 1) <= t Then
                    Exit Do
                End If
                data(k) = data(k - 1)
                k = k - 1
            Loop
            data(k) = t
        End If
    Next i
End Sub


