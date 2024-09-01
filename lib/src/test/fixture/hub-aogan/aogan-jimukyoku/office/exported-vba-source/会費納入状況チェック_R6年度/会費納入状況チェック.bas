Attribute VB_Name = "会費納入状況チェック"
Option Explicit

'本Subは、青森県眼科医会の会員が本年度の会費をすでに納めたかどうかを
'チェックすることを目的とするプログラムである。
'Visual Basic for Application言語によって実装されている。

'本Subは、青森県眼科医会の会員名簿Excelファイルと現金出納帳Excelファイルとの
'ふたつのExcelファイルの内容を照らし合わせることで必要な情報を導き出す。

'本ワークブックに「外部ファイルのパス」ワークシートがあって
'そのなかにふたつのExcelファイルのパスが書いてある。

'本Subは会員名簿のワークブックから本年度の会員一覧のワークシートをコピーして
'本ワークブックのなかに「work会員名簿」という名前のワークシートを作る。
'本Subは外部にあるワークブックを読むだけにして、書き換えない。
'本Subは各会員が会費を納付したかどうかの情報を「work会員名簿」ワークシートに書き込む。

'本Subは現金出納帳のExcelをREAD ONLYで参照する。現金出納帳にはいっさい書き込みしない。

'本モジュールを実行した結果としてある会員Mさんに×の印が付けられたとする。
'これだけでM会員が未納であると判断するのは危うい。
'もしも間違えてたら大変失礼だから注意せよ。
'人間が現金出納帳のデータをよく読んで確かめよう。
'現金出納帳が間違っているかもしれない。おおいにありうる。
'現金出納帳の間違いのせいで、じつはM会員が会費を振り込んでいたという情報を
'本モジュールを読み取ることができなかっただけかもしれない。
'会員である先生方に迷惑をかけないよう、十分に注意せよ。

Public Sub Main()

    Dim modName As String: modName = "会費納入状況チェック"
    Dim procName As String: procName = "Main"

    'イミディエイト・ウインドウを消す
    Call BbLog.Clear
    
    
    Call BbLog.Info(modName, procName, "会費納入状況チェックを行います")
    
    '会員名簿Excelファイルのパス
    Dim memberFile As String: memberFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B2")
    Call BbLog.Info(modName, procName, "会員名簿: " & memberFile)
    
    '現金出納帳Excelファイルのパス
    Dim cashbookFile As String: cashbookFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B3")
    Call BbLog.Info(modName, procName, "現金出納帳: " & cashbookFile)
    
    
    '============================================================================
    '外部にある会員名簿Excelファイルの "R6年度" シートをカレントのワークブックに
    'コピーする。"work会員名簿"シートが作られる。その内容をListObjectとして取り出す。
    Dim memberTable As ListObject
    Set memberTable = MbMemberTableUtil.FetchMemberTable(memberFile, "R6年度", ThisWorkbook)
    Call BbLog.Info(modName, procName, "会員の人数 memberTable.ListRows.Count=" & memberTable.ListRows.Count)
    
    'ListObjectの右端に列を追加する。列の名前を「会費納入状況」とする
    
    
    'work現金出納帳ワークシートを作ってCashbookオブジェクトを掴む
    Dim cb As Cashbook
    Set cb = OpenCashbook()
    Call BbLog.Info(modName, procName, "現金出納帳の行数 cb.Count=" & cb.Count)
    
    'チェックの対象とすべき開始日と終了日を指定したうえでCashSelectorオブジェクトを取得する
    Dim cs As CashSelector: Set cs = CbFactories.CreateCashSelector(cb, #4/1/2024#, #3/31/2025#)
    
    '各会員が会費を納入したかどうか調べてwork会員名簿に書き込む
    
    '会員名簿の全行についてループ
    Dim i As Long
    For i = 1 To memberTable.ListRows.Count
        '氏名漢字と氏名カナの２セルに字が書いてある行つまり名簿として有効な行を選ぶ
        Dim nameKanji As Variant: Set nameKanji = memberTable.ListColumns("氏名").DataBodyRange(i)
        Dim nameKana As Variant: Set nameKana = memberTable.ListColumns("氏名カナ").DataBodyRange(i)
        Dim entitlement As Variant: Set entitlement = memberTable.ListColumns("資格").DataBodyRange(i)
        If Not nameKanji.Value = "" And Not nameKana = "" Then
            'この人の氏名カナをキーとして現金出納帳を検索する
            Dim csList As CashList: Set csList = FindPaymentBy(cs, nameKana)
            '資格がAの人、Bの人、Cの人、Dの人について現金出納帳と照らし合わせる
            If entitlement = "A" Or entitlement = "B" Or entitlement = "C" Or entitlement = "D" Then
                If csList.Count = 1 Then
                    '通常どおり
                    Dim payedAt As String: payedAt = "R" & csList.Items(1).YY & "/" & csList.Items(1).MM & "/" & csList.Items(1).DD
                    Call PrintFinding(i, nameKana, entitlement, "◎ " & payedAt)
                    Call RecordFindingIntoMemberTable(memberTable, i, "◎ " & payedAt)
                ElseIf csList.Count > 1 Then
                    '同一名義人から2件以上の入金あり。おかしい。？を出力する。
                    Call PrintFinding(i, nameKana, entitlement, csList.Count & "?")
                    Call RecordFindingIntoMemberTable(memberTable, i, csList.Count & "?")
                    
                    'ひとりの会員が個人として振り込んだほかに勤務先医院が医院名義で当該会員の会費を振り込んだケースがあった。
                    '別々の名義人からの振込だから、このプログラムが重複を検出することはできなかった。
                    '出納事務を担当する者は未納を検出するだけでなく、目を凝らして重複納入も検出する必要がある。
                    
                Else
                    'ゼロなら会費が未納の可能性あり。×を出力する。
                    Call PrintFinding(i, nameKana, entitlement, "×")
                    Call RecordFindingIntoMemberTable(memberTable, i, "×")
                End If
            ElseIf entitlement Like "*弘大*" Then
                '資格がB弘大、C弘大については現金出納帳との照合をスキップする
                '弘大医局事務局が全員ぶんまとめて振り込むから個人単位の照合ができない。まあいいや、というわけで
                Call PrintFinding(i, nameKana, entitlement, "〇")
                Call RecordFindingIntoMemberTable(memberTable, i, "〇")
            Else
                '資格が免除、退会、その他の場合は現金出納帳をチェックせず△印を出力する
                Call PrintFinding(i, nameKana, entitlement, "△")
                Call RecordFindingIntoMemberTable(memberTable, i, "△")
            End If
        End If
    Next i
    Call BbLog.Info(modName, procName, "会費納入状況チェックを完了しました")
End Sub

'現金出納帳ワークシートに外部からデータをロードしてCashbokオブジェクトを返す
Private Function OpenCashbook() As Cashbook
    Dim wb As Workbook
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B3"))
    Dim sheetName As String: sheetName = "現金出納帳"
    Dim tableId As String: tableId = "CashbookTable1"
    Dim cb As Cashbook: Set cb = CbFactories.CreateCashbook(wb, sheetName, tableId)
    Set OpenCashbook = cb
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
End Function


Private Function FindPaymentBy(ByVal cs As CashSelector, ByVal nameKana As String)
    Dim csA, csB, csC, csD As CashList
    Set csA = cs.SelectCashListByMatchingDescription(AccountType.Income, "会費", "A会員", nameKana)
    Set csB = cs.SelectCashListByMatchingDescription(AccountType.Income, "会費", "B会員", nameKana)
    Set csC = cs.SelectCashListByMatchingDescription(AccountType.Income, "会費", "C会員", nameKana)
    Set csD = cs.SelectCashListByMatchingDescription(AccountType.Income, "会費", "D会員", nameKana)
    If csA.Count > 0 Then
        Set FindPaymentBy = csA
    ElseIf csB.Count > 0 Then
        Set FindPaymentBy = csB
    ElseIf csC.Count > 0 Then
        Set FindPaymentBy = csC
    ElseIf csD.Count > 0 Then
        Set FindPaymentBy = csD
    Else
        Set FindPaymentBy = CbFactories.CreateEmptyCashList()
    End If
End Function


Private Sub PrintFinding(ByVal i As Long, _
                            ByVal nameKana As String, _
                            ByVal entitlement As String, _
                            ByVal status As String)
    Dim message As String
    message = "|" & i & "|" & nameKana & "|" & entitlement & "|" & status & "|"
    Call BbLog.Info("会費納入状況チェック", "Main", message)
End Sub
                    


Private Sub RecordFindingIntoMemberTable(ByVal memberTable As ListObject, _
                                    ByVal i As Long, _
                                    ByVal status As String)
    '会員名簿テーブルのi番目の行に判定結果を書き込む。
    memberTable.ListColumns("会費納入状況").DataBodyRange(i).Value = status
    'statusが未納ならばその行の文字を赤色に変更する
    If status Like "×" Then
        With memberTable.ListColumns("氏名").DataBodyRange(i).Font
            .Color = RGB(255, 64, 64)
            .Underline = True
        End With
    End If
End Sub

