Attribute VB_Name = "BbArraySort"
Option Explicit

' 配列をソートする

' https://www.tipsfound.com/vba/02020

Public Sub Test_InsertionSort()
    Call BbLog.Clear
    Dim data() As Variant
    data = Array(7, 2, 6, 3, 9, 1, 8, 0, 5, 4)
    ' 並べ替える
    Call InsertionSort(data, LBound(data), UBound(data))
    Dim d As Variant
    For Each d In data
        Debug.Print d
        ' 0 1 2 3 4 5 6 7 8 9
    Next
End Sub


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


