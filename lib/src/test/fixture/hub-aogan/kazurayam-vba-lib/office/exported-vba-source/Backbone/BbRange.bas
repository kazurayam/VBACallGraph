Attribute VB_Name = "BbRange"
Option Explicit

' BbRangeモジュール　Rangeオブジェクトを操作するヘルパを実装する

'Rangeオブジェクトを受け取り、内容文字列を配列に変換して返す。ただし
'返される配列の内容はStringで、重複が無い。
Public Function GetUniqueItems(r As Range) As Variant
    Dim d As Dictionary: Set d = New Dictionary
    Dim item As Variant
    Dim k As String
    For Each item In r
        k = CStr(item)   'キーを明示的にStringにする
        If d.Exists(k) = False Then
            d.Add k, item
            'Debug.Print "add " & k
        End If
    Next
    GetUniqueItems = d.keys
End Function

