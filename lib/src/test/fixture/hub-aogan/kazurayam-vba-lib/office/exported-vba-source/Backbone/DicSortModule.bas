Attribute VB_Name = "DicSortModule"
Option Explicit


' DicSortModule: Dictionary連想配列をキーの昇順にソートするSub DicSortを提供する
' DictionaryのキーがStringで、valueもStringであることを前提する。
' valueがオブジェクト型であるようなDictionaryは扱えない。型が不一致（Stringでないから）のエラーが発生する。

' Qiitaに掲載された記事 [VBAでDictionary（連想配列）を辞書順にソートする](https://qiita.com/daik/items/682743bb8bcd8b5f0689)が
' 公開したコードをまるまるコピーした。

' Dictionaryを参照引数にし、これをソートする破壊的プロシージャ。
Public Sub DicSort(ByRef dic As Object)
    Dim i As Long, j As Long
    Dim varTmp() As String
    Dim key As Variant
    Dim dicSize As Long: dicSize = dic.Count
    
    'dicのサイズに合わせて２次元配列のサイズを調整して
    ReDim varTmp(dicSize + 1, 2)

    ' Dictionaryが空であるか、サイズが1以下であればソート不要
    If dic Is Nothing Or dicSize < 2 Then
        Exit Sub
    End If

    ' Dictionaryのキーとvalueを２次元配列に転写
    i = 0
    For Each key In dic
        varTmp(i, 0) = key
        varTmp(i, 1) = dic(key)
        i = i + 1
    Next
    
    '２次元配列をキーの昇順でソート
    Call QuickSort(varTmp, 0, dicSize - 1)
    dic.RemoveAll

    '２次元配列の内容をDictionaryに上書きする
    For i = 0 To dicSize - 1
        dic(varTmp(i, 0)) = varTmp(i, 1)
    Next
    
End Sub

'' String型で2列の二次元配列を受け取り、これの1列目でクイックソートする（ほんとはCompareメソッドを渡すAdapterパターンで書きたいところ、VBAのオブジェクト指向厳しい感じで妥協）
Private Sub QuickSort(ByRef targetVar() As String, ByVal min As Long, ByVal max As Long)
    Dim i, j As Long
    Dim tmp As String
    Dim pivot As Variant

    If min < max Then
        i = min
        j = max
        pivot = StrMed3(targetVar(i, 0), targetVar(Int((i + j) / 2), 0), targetVar(j, 0))
        Do
            Do While StrComp(targetVar(i, 0), pivot) < 0
                i = i + 1
            Loop
            Do While StrComp(pivot, targetVar(j, 0)) < 0
                j = j - 1
            Loop
            If i >= j Then Exit Do

            tmp = targetVar(i, 0)
            targetVar(i, 0) = targetVar(j, 0)
            targetVar(j, 0) = tmp

            tmp = targetVar(i, 1)
            targetVar(i, 1) = targetVar(j, 1)
            targetVar(j, 1) = tmp

            i = i + 1
            j = j - 1

        Loop
        Call QuickSort(targetVar, min, i - 1)
        Call QuickSort(targetVar, j + 1, max)

    End If
End Sub

'' String型のx, y, z を辞書順比較し二番目のものを返す
Private Function StrMed3(ByVal x As String, ByVal y As String, ByVal z As String)
    If StrComp(x, y) < 0 Then
        If StrComp(y, z) < 0 Then
            StrMed3 = y
        ElseIf StrComp(z, x) < 0 Then
            StrMed3 = x
        Else
            StrMed3 = z
        End If
    Else
        If StrComp(z, y) < 0 Then
            StrMed3 = y
        ElseIf StrComp(x, z) < 0 Then
            StrMed3 = x
        Else
            StrMed3 = z
        End If
    End If
End Function

'' テストメソッド
Sub TestDicSort()

    Dim output As String
    Dim dic As Object
    Dim key As Variant

    Set dic = CreateObject("Scripting.Dictionary")

    dic("g") = "gggg"
    dic("9") = "999"
    dic("を") = "をををを"
    dic("4") = "444"
    dic("あ") = "ああああ"
    dic("(") = "(((("
    dic("a") = "aaaa"

    output = "##before" & vbNewLine

    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key

    Call DicSort(dic)

    output = output + vbNewLine & "##after" & vbNewLine
    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key

    Debug.Print output
End Sub

