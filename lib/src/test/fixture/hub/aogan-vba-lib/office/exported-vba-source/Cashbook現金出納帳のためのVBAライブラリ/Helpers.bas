Attribute VB_Name = "Helpers"
Option Explicit

' Cashbookに記載された入出金データを勘定科目ごとに分類し件数を数えて表示する
Public Sub PrintAccounts(ByVal cb As Cashbook, ByVal rpUnit As String, ByVal positiveLike As Boolean)
    'Debug.Print "PrintAccounts was called"
    Dim accFinder As AccountsFinder: Set accFinder = CreateAccountsFinder(cb)
    Call accFinder.Initialize(cb, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    'Debug.Print "initialized AccountFinder"
    Dim dic As Dictionary: Set dic = accFinder.FindAccounts(rpUnit, positiveLike)
    Debug.Print "dic was created"
    Debug.Print "dic.Count=" & dic.Count
    ' dicのキーはString、たとえば「支出/事務費/通信費」
    ' dicのvalueはCashListオブジェクト
    
    ' キーの一覧を配列にとりだしてソートする
    Dim var As Variant
    var = dic.Keys
    Call InsertionSort(var, LBound(var), UBound(var))
    
    ' キーごとに明細をPrintする
    Dim cList As CashList
    Dim msg As String
    Dim key As Variant
    For Each key In var
        Set cList = dic(key)
        Debug.Print key,
        Debug.Print cList.Count & "件",
        If key Like "収入/*" Then
            Debug.Print cList.SumOfIncomeAmount() & " 円"
        ElseIf key Like "支出/*" Then
            Debug.Print cList.SumOfExpenseAmount() & " 円"
        Else
            Debug.Print ""
        End If
    Next
End Sub
