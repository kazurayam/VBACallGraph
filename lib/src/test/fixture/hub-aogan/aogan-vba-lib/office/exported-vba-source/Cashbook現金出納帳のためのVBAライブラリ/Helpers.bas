Attribute VB_Name = "Helpers"
Option Explicit

' Helpers

' Cashbookに記載された入出金データを勘定科目ごとに分類し件数を数えて表示する
Public Sub PrintAccounts(ByVal cb As Cashbook, ByVal rpUnit As String, ByVal positiveLike As Boolean)
    'Debug.Print "PrintAccounts was called"
    Dim cbTransformer As CashbookTransformer
    Set cbTransformer = Factories.CreateCashbookTransformer(cb)
    Call cbTransformer.Initialize(cb, periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#)
    'Debug.Print "initialized AccountFinder"
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts(rpUnit, positiveLike)
    Debug.Print "cashListDic.Count=" & cashListDic.Count
    ' dicのキーはString、たとえば「支出/事務費/通信費」
    ' dicのvalueはCashListオブジェクト
    
    ' キーの一覧を配列にとりだしてソートする
    Dim accounts As Variant
    accounts = cashListDic.Keys
    Call BbArraySort.InsertionSort(accounts, LBound(accounts), UBound(accounts))
    
    ' キーごとに明細をPrintする
    Dim account_ As Variant
    Dim cashList_ As CashList
    Dim msg As String
    For Each account_ In accounts
        Set cashList_ = cashListDic(account_)
        Debug.Print account_,
        Debug.Print cashList_.Count & "件",
        If account_ Like "収入/*" Then
            Debug.Print cashList_.SumOfIncomeAmount() & " 円"
        ElseIf account_ Like "支出/*" Then
            Debug.Print cashList_.SumOfExpenseAmount() & " 円"
        Else
            Debug.Print ""
        End If
    Next
End Sub
