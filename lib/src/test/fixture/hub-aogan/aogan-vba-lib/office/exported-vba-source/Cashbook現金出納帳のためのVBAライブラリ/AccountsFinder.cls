VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccountsFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' AccountsFinderクラス

' Cashbookの内容を走査し,ある期間において入出金があった勘定科目Accountを特定し、Dictionaryを返す。
' DictionaryのitemのキーはAccountを表現する文字列（account.itsAccout）である。たとえば "支出/事務費/通信費"
' DictionaryのitemのvalueはCashListオブジェクトである。
' Dictionaryのキーの一覧を調べればある期間において少なくとも１件の入出金があった勘定科目を知ることができる。
' ひとつのキーつまりAccountを指定してDictionaryからitemをgetすればCashListが得られる。
' CashListのCount()をみれば入出金の件数がわかる。
' CashListのSumOfIncomeAmount()およびSumOfExpenseAmount()をみれば勘定科目ごとの合計金額がわかる。

' Dictionaryを使う。そのためにMicrosoft Scripting Runtimeを参照する
' https://tonari-it.com/excel-vba-dictionary-object/

Private cb_ As Cashbook
Private periodStart_ As Date
Private periodEnd_ As Date
Private cSel_ As CashSelector
'


Public Sub Initialize(ByVal source As Cashbook, Optional ByVal periodStart As Date = #4/1/2022#, Optional ByVal periodEnd As Date = #3/31/2023#)
    Set cb_ = source
    periodStart_ = periodStart
    periodEnd_ = periodEnd
    Set cSel_ = New CashSelector
    Call cSel_.Initialize(cb_, periodStart_, periodEnd_)
End Sub

Public Function FindAccounts(Optional ofReportingUnit = "*", Optional positiveLike = True) As Dictionary
    ' Cashbookをスキャンしてキーを発見する
    Dim dic1 As Dictionary
    Set dic1 = New Dictionary
    Dim i As Long
    
    For i = 1 To cb_.Count
        Dim cs As Cash: Set cs = cb_.Items(i)
        '指定された年度のデータだけを選ぶ、デフォルトとして本年度のデータを選ぶ
        If periodStart_ <= cs.ToDate() And cs.ToDate() <= periodEnd_ Then
            ' 収支報告単位がパラメータと合致するCashだけを選ぶ。ただし
                'positiveLikeが True ならふつうにLikeｋ演算子を使う
                'positiveLikeが False ならNot Like演算子を使う
            If (positiveLike And cs.ReportingUnit Like ofReportingUnit) Or _
               (Not positiveLike And Not cs.ReportingUnit Like ofReportingUnit) Then

                Dim key As Variant
                key = cs.ItsAccount.ToString()
                
                If dic1.Exists(key) = False Then
                    dic1.Add key, CLng(1)  ' このkeyの出現回数として1を書き込む
                Else
                    Dim Count As Long: Count = dic1.item(key)
                    dic1(key) = CLng(Count + 1) ' このkeyの出現回数を+1する
                End If
            End If
        End If
    Next i
    
    ' キーとそれに関連するCashListオブジェクトの組を作ってDictionaryを構築する
    Dim dic2 As Dictionary: Set dic2 = New Dictionary
    Dim k As Variant
    
    For Each k In dic1.Keys
        Dim acc As Account: Set acc = New Account
        Call acc.Of(k)
        Set dic2(k) = cSel_.SelectCashList(acc.AccType, acc.AccountName, acc.SubAccountName, _
                                           ofReportingUnit, positiveLike)
    Next
    
    'ソート
    'Call DicSort(dic2)
    'しない。keyとvalueがともにStringであるDictionaryではソートできたが、
    'CashListオブジェクトをvalueに設定したらソートを実装するのが無理だった。
    'CallerはreturnされたDictionaryのKeysをソートして、ソート後の順序にしたがってDictionaryの内容を参照すべき
    
    Set FindAccounts = dic2
End Function


Public Function FindKeysAsString(Optional ofReportingUnit = "*") As String
    'FindAccounts関数が返すDictionaryオブジェクトのKeyを取り出して
    'Keyを1行として、複数個のKeyをひとつに連結したStringを返す。
    'デバッグ用のショートカット
    Dim dic As Dictionary: Set dic = FindAccounts(ofReportingUnit)
    Dim arr
    arr = dic.Keys
    Dim s As String
    s = ""
    Dim i As Long
    For i = 0 To dic.Count - 1   'VBAの配列は0始まりだよん、Collectionのインデックスは1始まりだが
        s = s & arr(i) & vbNewLine
    Next i
    FindKeysAsString = s
End Function
