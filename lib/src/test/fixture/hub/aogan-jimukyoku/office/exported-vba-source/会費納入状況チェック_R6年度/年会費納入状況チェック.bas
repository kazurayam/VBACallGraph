Attribute VB_Name = "年会費納入状況チェック"
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

    'イミディエイト・ウインドウを消す
    Call KzCls
    
    Debug.Print ("会費納入状況チェックを行います")
    
    '会員名簿Excelファイルのパス
    Dim memberFile As String: memberFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B2")
    Debug.Print ("会員名簿: " & memberFile)
    
    '現金出納帳Excelファイルのパス
    Dim cashbookFile As String: cashbookFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B3")
    Debug.Print ("現金出納帳: " & cashbookFile)
    
    
    '============================================================================
    '外部にある会員名簿Excelファイルの "R6年度" シートをカレントのワークブックに
    'コピーする。"work会員名簿"シートが作られる。その内容をListObjectとして取り出す。
    Dim memberTable As ListObject
    Set memberTable = AoMemberUtils.FetchMemberTable(memberFile, "R6年度", ThisWorkbook)
    Set memberTable = OpenMemberTable("R6年度", True)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.Count
    
    'work現金出納帳ワークシートを作ってCashbookオブジェクトを掴む
    Dim cb As Cashbook
    Set cb = OpenCashbook()
    Debug.Print "cb.Count=" & cb.Count
    Dim cs As CashSelector: Set cs = CreateCashSelector(cb, #4/1/2023#, #3/31/2024#)
    
    '各会員が会費を納入したかどうか調べてwork会員名簿に書き込む
    Dim i As Long
    '会員名簿の全行についてループ
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
    
End Sub

Private Function OpenMemberTable(Optional ByVal sheetName As String = "R6年度", Optional ByVal renew As Boolean = False) As ListObject
    'マスタの会員名簿ワークブックを開いて名簿のワークシートをカレントのワークブックに取り込む
    Dim targetWorkbook As Workbook: Set targetWorkbook = ThisWorkbook
    Dim targetSheetName As String: targetSheetName = "work会員名簿"
    'work会員名簿ワークシートがまだ存在しない、または
    'ワークシートがすでに存在するがrenewパラメータがTrueならば
    If (Not KzVerifyWorksheetExists(targetSheetName)) Or _
        (KzVerifyWorksheetExists(targetSheetName) And renew = True) Then
        Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(mbGetPathOfAoganMembers())
        Dim sourceSheetName As String: sourceSheetName = sheetName
        Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
        '会員名簿ワークブックを閉じる
        Application.DisplayAlerts = False   '「変更内容を保存しますか」ダイアログを表示しない
        sourceWorkbook.Close
        Application.DisplayAlerts = True
    End If
    'カレントのワークブックに取り込んだワークシートのなかのテーブルを掴む
    Dim ws As Worksheet: Set ws = targetWorkbook.Worksheets(targetSheetName)
    Dim tbl As ListObject: Set tbl = ws.ListObjects("MembersTable13")
    
    '会員名簿のテーブルの右端列が「列4」という名前であるところを「会費納入状況」に変更する
    tbl.ListColumns(tbl.ListColumns.Count).Name = "会費納入状況"
    
    '列幅を調整
    tbl.ListColumns("氏名カナ").Range.EntireColumn.AutoFit
    tbl.ListColumns("勤務先名").Range.EntireColumn.AutoFit
    tbl.ListColumns("異動").Range.EntireColumn.AutoFit
    tbl.ListColumns("会費納入状況").Range.EntireColumn.AutoFit
    
    '重要でない列を非表示にする。シートを印刷したときに便利なように
    tbl.ListColumns("年齢").Range.EntireColumn.Hidden = True
    tbl.ListColumns("医登No").Range.EntireColumn.Hidden = True
    tbl.ListColumns("医登録日").Range.EntireColumn.Hidden = True
    tbl.ListColumns("日眼医登録").Range.EntireColumn.Hidden = True
    tbl.ListColumns("勤務〒").Range.EntireColumn.Hidden = True
    tbl.ListColumns("勤務先住所").Range.EntireColumn.Hidden = True
    tbl.ListColumns("勤先TELNo").Range.EntireColumn.Hidden = True
    tbl.ListColumns("自宅〒").Range.EntireColumn.Hidden = True
    tbl.ListColumns("自宅住所").Range.EntireColumn.Hidden = True
    tbl.ListColumns("自宅TELNo").Range.EntireColumn.Hidden = True
    tbl.ListColumns("携帯番号").Range.EntireColumn.Hidden = True
    
    Set OpenMemberTable = tbl
End Function

Private Function OpenCashbook() As Cashbook
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Dim sheetName As String: sheetName = "現金出納帳"
    Dim tableId As String: tableId = "CashbookTable1"
    Dim cb As Cashbook: Set cb = CreateCashbook(wb, sheetName, tableId)
    Set OpenCashbook = cb
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
End Function

'Private Function OpenCashbook(Optional ByVal renew As Boolean = False) As Cashbook
'    'マスターの現金出納帳ワークブックをひらいて現金出納帳のワークシートをカレントのワークブックに取り込む
'    Dim targetWorkbook As Workbook: Set targetWorkbook = ThisWorkbook
'    Dim targetSheetName As String: targetSheetName = "work現金出納帳"
'    'work会員名簿ワークシートがまだ存在しない、または
'    'ワークシートがすでに存在するがrenewパラメータがTrueならば
'    If (Not KzVerifyWorksheetExists(targetSheetName)) Or _
'        (KzVerifyWorksheetExists(targetSheetName) And renew = True) Then
'        Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(GetPathOfAoganCashbook)
'        Dim sourceSheetName As String: sourceSheetName = "現金出納帳"
'        Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
'        '現金出納帳ワークブックを閉じる
'        Application.DisplayAlerts = False   '「変更内容を保存しますか」ダイアログを表示しない
'        sourceWorkbook.Close
'        Application.DisplayAlerts = True
'    End If
'    'カレントのワークブックに取り込んだワークシートのなかのテーブルを掴む
'    Dim cb As Cashbook: Set cb = CreateCashbook(targetWorkbook, targetSheetName)
'    '
'    Set OpenCashbook = cb
'End Function



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
        Set FindPaymentBy = CreateEmptyCashList()
    End If
End Function


Private Sub PrintFinding(ByVal i As Long, _
                            ByVal nameKana As String, _
                            ByVal entitlement As String, _
                            ByVal status As String)
    Debug.Print "|"; i & "|" & nameKana & "|" & entitlement & "|" & status & "|"
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

