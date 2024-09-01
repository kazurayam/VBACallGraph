VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' CashSelectorクラス

' 指定された条件に従ってCashbookのなから条件にかなうCashの集合を選択しCashListにして返す

Private cb_ As Cashbook
Private periodStart_ As Date
Private periodEnd_ As Date


Public Sub Initialize(ByVal source As Cashbook, Optional ByVal periodStart As Date = #4/1/2022#, Optional ByVal periodEnd As Date = #3/31/2023#)
    Set cb_ = source
    periodStart_ = periodStart
    periodEnd_ = periodEnd
End Sub

Public Function SelectCashListByAccount(ByVal acc As account, _
                                        Optional ofReportingUnit = "*", Optional positiveLike = True) As CashList
    Dim result As CashList
    Set result = SelectCashList(acc.AccType(), acc.AccountName(), acc.SubAccountName(), _
                            ofReportingUnit, positiveLike)
    Set SelectCashListByAccount = result
End Function


Public Function SelectCashList(ByVal ofAccountType As AccountType, ByVal ofAccountName, _
                               ByVal ofSubAccountName, _
                               Optional ofReportingUnit = "*", Optional positiveLike = True) As CashList
    Dim result As CashList: Set result = New CashList
    Dim i As Long
    For i = 1 To cb_.Count
        Dim cs As cash: Set cs = cb_.Items(i)
        Dim csDate As Date: csDate = cs.ToDate()
        '指定された年度のデータだけを選ぶ、デフォルトとして本年度のデータを選ぶ
        If periodStart_ <= csDate And csDate <= periodEnd_ Then
            ' positiveLikeがtrueならCashの収支報告単位がパラメータと合致するCashデータを選ぶ
            ' positiveLikeがfalseならCashの収支報告単位がパラメータと合致しないCashデータを選ぶ
            If (positiveLike And cs.ReportingUnit Like ofReportingUnit) Or _
               (Not positiveLike And Not cs.ReportingUnit Like ofReportingUnit) Then
                If cs.ItsAccount.AccType = ofAccountType Then
                    ' Cashの科目と補助科目がパラメータと合致したら
                    If cs.ItsAccount.AccountName Like ofAccountName And _
                        cs.ItsAccount.SubAccountName Like ofSubAccountName Then
                        'つかまえる
                        result.Add cs
                    End If
                End If
            End If
        End If
    Next i
    Set SelectCashList = result
End Function


'収入支出区分＋科目＋補助科目に加え、摘要の値のなかにkeyと合致する部分があるかどうかを調べる。
'合致するCashオブジェクトをリストにして返す。
'keyと摘要の値を比較するにあたっては単純化のため半角スペースと全角スペースを除去してから比較する。
Public Function SelectCashListByMatchingDescription( _
        ByVal ofAccountType As AccountType, _
        ByVal ofAccountName, _
        ByVal ofSubAccountName, _
        ByVal againstKey As String, _
        Optional ofReportingUnit = "*") As CashList
    Dim result As CashList: Set result = New CashList
    Dim i As Long
    For i = 1 To cb_.Count
        Dim cs As cash: Set cs = cb_.Items(i)
        Dim csDate As Date: csDate = cs.ToDate()
        '指定された年度のデータだけを選ぶ、デフォルトとして本年度のデータを選ぶ
        If periodStart_ <= csDate And csDate <= periodEnd_ Then
            ' Cashの収支報告単位がパラメータと合致するCashデータを選ぶ
            If cs.ReportingUnit Like ofReportingUnit Then
                
                If cs.ItsAccount.AccType = ofAccountType Then
                    ' Cashの科目と補助科目がパラメータと合致したら
                    If cs.ItsAccount.AccountName Like ofAccountName And _
                        cs.ItsAccount.SubAccountName Like ofSubAccountName Then
                        '摘要の値がどこかにkeyを含んでいたら
                        If InStr(TrimAllSpaces(cs.Description), TrimAllSpaces(againstKey)) <> 0 Then
                            'つかまえる
                            result.Add cs
                            
                        End If
                    End If
                End If
            End If
        End If
    Next i
    Set SelectCashListByMatchingDescription = result
End Function

Private Function TrimAllSpaces(ByVal str As String) As String
    Dim s As String
    s = Replace(str, " ", "")   '半角スペースをすべて削除
    s = Replace(s, "　", "")    '全角スペースをすべて削除
    TrimAllSpaces = s
End Function
