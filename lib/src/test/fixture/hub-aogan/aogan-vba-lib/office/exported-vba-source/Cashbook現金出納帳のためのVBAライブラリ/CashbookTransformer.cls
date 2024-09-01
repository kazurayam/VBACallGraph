VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashbookTransformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' CashbookTransformerクラス
' Cashbookの内容を走査し、指定された開始から終了までの期間において入出金があった
' 勘定科目Accountを特定し、Dictionaryを返す。
' DictionaryのItemのキーはAccountすなわち勘定科目を表現する文字列である。
' たとえば "支出/事務費/通信費" のような文字列がキーとなる。
' DictionaryのItemの値はCashListオブジェクトである。
' Dictionaryのキーの一覧を調べれば当該期間において少なくとも１件の入出金があった
' 勘定科目のリストを得ることができる。
' ひとつのキーつまりAccountを指定してDictionaryからCashListが取得できる。
' それは勘定科目毎の入出金記録の明細にほかならない。
' CashListオブジェクトのCount()をみれば入出金の件数がわかる。
' CashListのSumOfIncomeAmount()およびSumOfExpenseAmount()をみれば勘定科目ごとの
' 入出金の小計額がわかる。

Private cashbook_ As Cashbook
Private periodStart_ As Date
Private periodEnd_ As Date
Private cashSelector_ As CashSelector
'

Public Sub Initialize(ByVal source As Cashbook, _
                    Optional ByVal periodStart As Date = #4/1/2022#, _
                    Optional ByVal periodEnd As Date = #3/31/2023#)
    Set cashbook_ = source
    periodStart_ = periodStart
    periodEnd_ = periodEnd
    Set cashSelector_ = New CashSelector
    Call cashSelector_.Initialize(cashbook_, periodStart_, periodEnd_)
End Sub

' 勘定科目をキーとし、CashListを値とするDictionaryを構築して返す
Public Function ByAccounts(Optional ofReportingUnit = "*", Optional positiveLike = True) As Dictionary
    ' Cashbookをスキャンしてキーを発見する
    Dim dic1 As Dictionary: Set dic1 = New Dictionary
    Dim i As Long
    
    For i = 1 To cashbook_.Count
        Dim cash_ As cash: Set cash_ = cashbook_.Items(i)
        '指定された年度のデータだけを選ぶ、デフォルトとして本年度のデータを選ぶ
        If periodStart_ <= cash_.ToDate() And cash_.ToDate() <= periodEnd_ Then
            ' 収支報告単位がパラメータと合致するCashだけを選ぶ。ただし
                'positiveLikeが True ならふつうにLikeｋ演算子を使う
                'positiveLikeが False ならNot Like演算子を使う
            If (positiveLike And cash_.ReportingUnit Like ofReportingUnit) Or _
               (Not positiveLike And Not cash_.ReportingUnit Like ofReportingUnit) Then

                Dim key As Variant
                key = cash_.ItsAccount.ToString()  ' keyは"支出/事務費/通信費"のような文字列
                
                If dic1.Exists(key) = False Then
                    dic1.Add key, CLng(1)  ' このkeyの出現回数として1を書き込む
                Else
                    Dim count_ As Long: count_ = dic1(key)
                    dic1(key) = CLng(count_ + 1) ' このkeyの出現回数を+1する
                End If
            End If
        End If
    Next i
    
    ' キーとそれに関連するCashListオブジェクトの組を作ってDictionaryを構築する
    Dim dic2 As Dictionary: Set dic2 = New Dictionary
    Dim k As Variant
    
    For Each k In dic1.Keys
        Dim account_ As account: Set account_ = New account
        Call account_.Of(k)
        Set dic2(k) = cashSelector_.SelectCashList( _
                            account_.AccType, _
                            account_.AccountName, _
                            account_.SubAccountName, _
                            ofReportingUnit, positiveLike)
    Next
    
    'キーの昇順にソート済みの結果を返したいところだが、しない。
    'というのも実現が難しかったから。
    'keyとvalueがともにStringであるDictionaryではソートできたが、
    'CashListオブジェクトをvalueに設定したらソートを実装するのが無理だった。
    'CallerはreturnされたDictionaryのKeysをソートして、ソート後の順序にしたがって
    'Dictionaryの内容を参照すべき
    
    Set ByAccounts = dic2
    
End Function


Public Function FindKeysAsString(Optional ofReportingUnit = "*") As String
    'FindAccounts関数が返すDictionaryオブジェクトのKeyを取り出して
    'Keyを1行として、複数個のKeyをひとつに連結したStringを返す。
    'デバッグ用のショートカット
    Dim cashListDic As Dictionary: Set cashListDic = ByAccounts(ofReportingUnit)
    Dim accounts As Variant: accounts = cashListDic.Keys
    'キーを昇順にソートする
    Call BbArraySort.InsertionSort(accounts, LBound(accounts), UBound(accounts))
    
    Dim s As String
    s = ""
    Dim i As Long
    For i = 0 To cashListDic.Count - 1   'VBAの配列は0始まりだよん、Collectionのインデックスは1始まりだが
        s = s & accounts(i) & vbNewLine
    Next i
    FindKeysAsString = s
End Function
