Attribute VB_Name = "WriteSettlement"

Option Explicit

'AccountSumモジュール

'現金出納帳ワークブックのデータを入力として【現金出納記録】シートを作成する。
'【現金出納記録】シートは 勘定科目ごとの小計 Sum of account を算出する。
'勘定科目ごとの小計を問い合わせることのできる Get小計 Functionを提供する。
'決算書シートなどほかのシートのセルが参照式で Get小計 関数を使えば
'勘定科目ごとの小計を参照することができる。

Const TheSheetName As String = "現金出納記録"
Const RecTableName As String = "テーブル現金出納記録"
Const SumTableName As String = "テーブル勘定科目ごとの小計"

Enum enumSource
    領収書番号 = 1
    令和
    月
    日
    収入科目
    収入補助科目
    支出科目
    支出補助科目
    収支報告単位
    摘要
    借方金額
    貸方金額
    差引残高
    立替者
    清算済みか
End Enum

Enum enumTarget
    勘定科目 = 1
    年月日
    収入科目
    収入補助科目
    支出科目
    支出補助科目
    収支報告単位
    摘要
    借方金額
    貸方金額
End Enum


'==============================================================================

Public Sub 入出金記録を取り込む(ByVal wsCashSource As Worksheet, _
                                ByVal wsCashTarget As Worksheet, _
                                Optional ByVal cashbookTableName As String = "CashbookTable1", _
                                Optional ByVal periodStart As Date = #4/1/2022#, _
                                Optional ByVal periodEnd As Date = #3/31/2023#, _
                                Optional ofReportingUnit = "*", _
                                Optional positiveLike = True, _
                                Optional initializeTarget = True)
    Dim tblSource As ListObject
    Set tblSource = wsCashSource.ListObjects(cashbookTableName)
    Dim tblTarget As ListObject
    Set tblTarget = wsCashTarget.ListObjects(RecTableName)
     
    'initializeTargetオプションがTrueならば　ターゲットのテーブルを初期化する
    If initializeTarget Then
        tblTarget.DataBodyRange.Delete
    End If
    
    'ソースのテーブルからターゲットのテーブルに行を転写する。
    '【令和】+【月】+【日】を【年月日】に変換したり
    '不要な列を無視したりしながら。
    Dim i As Long
    For i = 1 To tblSource.ListRows.Count
        Dim rowSource As ListRow: Set rowSource = tblSource.ListRows(i)
        If Is取り込むべき(rowSource, periodStart, periodEnd, ofReportingUnit, positiveLike) Then
            'データを写す
            Call TransferRow(rowSource, tblTarget)
        End If
    Next i
    'A列（勘定科目）の幅を指定する
    wsCashTarget.Columns("A").ColumnWidth = 22
    'B列（年月日）の幅を適正にする
    wsCashTarget.Columns("B:B").AutoFit
End Sub


'rowSourceデータ１行をtblTarget標的であるテーブルへAddする
Public Sub TransferRow(rowSource As ListRow, tblTarget As ListObject)
    With tblTarget.ListRows.Add
        .Range(enumTarget.年月日).value = ToDate( _
                          rowSource.Range(enumSource.令和).value, _
                            rowSource.Range(enumSource.月).value, _
                            rowSource.Range(enumSource.日).value)
        .Range(enumTarget.収入科目).value = rowSource.Range(enumSource.収入科目).value
        .Range(enumTarget.収入補助科目).value = rowSource.Range(enumSource.収入補助科目).value
        .Range(enumTarget.支出科目).value = rowSource.Range(enumSource.支出科目).value
        .Range(enumTarget.支出補助科目).value = rowSource.Range(enumSource.支出補助科目).value
        .Range(enumTarget.収支報告単位).value = rowSource.Range(enumSource.収支報告単位).value
        .Range(enumTarget.勘定科目).value = To勘定科目(rowSource)
        .Range(enumTarget.摘要).value = rowSource.Range(enumSource.摘要).value
        .Range(enumTarget.借方金額).value = rowSource.Range(enumSource.借方金額).value
        .Range(enumTarget.貸方金額).value = rowSource.Range(enumSource.貸方金額).value
        
        '金額のセルの書式を設定する（通貨、3桁毎カンマで区切る、\無し）
        .Range(enumTarget.借方金額).NumberFormatLocal = "#,###"
        .Range(enumTarget.貸方金額).NumberFormatLocal = "#,###"
    End With
End Sub


Public Function Is取り込むべき(ByVal rowSource As ListRow, _
                                Optional ByVal periodStart As Date = #4/1/2022#, _
                                Optional ByVal periodEnd As Date = #3/31/2023#, _
                                Optional ofReportingUnit = "*", _
                                Optional positiveLike = True) As Boolean
    If Is金額が非ゼロだ(rowSource) Then
        If Is年度内だ(rowSource, periodStart, periodEnd) Then
            If Is収支報告単位が該当する(rowSource, ofReportingUnit, positiveLike) Then
                Is取り込むべき = True
            Else
                Is取り込むべき = False
            End If
        Else
            Is取り込むべき = False
        End If
    Else
        Is取り込むべき = False
    End If
End Function


Public Function Is金額が非ゼロだ(ByVal rowSource As ListRow) As Boolean
    If rowSource.Range(enumSource.貸方金額).value <> 0 Or _
        rowSource.Range(enumSource.借方金額).value <> 0 Then
        Is金額が非ゼロだ = True
    Else
        Is金額が非ゼロだ = False
    End If
End Function


Public Function Is年度内だ(ByVal rowSource As ListRow, _
                            Optional ByVal periodStart As Date = #4/1/2022#, _
                            Optional ByVal periodEnd As Date = #3/31/2023#) As Boolean
    Dim dt As Date
    dt = ToDate(rowSource.Range(enumSource.令和).value, _
                rowSource.Range(enumSource.月).value, _
                rowSource.Range(enumSource.日).value)
    If periodStart <= dt And dt <= periodEnd Then
        Is年度内だ = True
    Else
        Is年度内だ = False
    End If
End Function


Public Function Is収支報告単位が該当する(ByVal rowSource As ListRow, _
                                         Optional ofReportingUnit = "*", _
                                         Optional positiveLike = True) As Boolean
    Dim rpUnit As String: rpUnit = rowSource.Range(enumSource.収支報告単位)
    If (positiveLike And rpUnit Like ofReportingUnit) Or _
        (Not positiveLike And Not rpUnit Like ofReportingUnit) Then
        Is収支報告単位が該当する = True
    Else
        Is収支報告単位が該当する = False
    End If
End Function


'==============================================================================


'入力行の【収入科目】【収入補助科目】【支出科目】【支出補助科目】から
'  "事務費/通信費"
'  "会費/A会員"
'のような文字列を合成して返す。これを『入出金記録』シートの【勘定科目】に利用する。
'【勘定科目】と【年月日】をキーとしてテーブルの行をソートするために。
'VBAのテーブルのSort機能はキーを3個までしか指定できないから、【勘定科目】を合成することが必要だった
Public Function To勘定科目(ByVal rowSource As ListRow) As String
    If rowSource.Range(enumSource.支出科目).value <> "" And _
        rowSource.Range(enumSource.支出補助科目).value <> "" Then
        
        To勘定科目 = "支出" & "/" _
                        & rowSource.Range(enumSource.支出科目).value _
                        & "/" & rowSource.Range(enumSource.支出補助科目).value
    
    ElseIf rowSource.Range(enumSource.収入科目).value <> "" And _
            rowSource.Range(enumSource.収入補助科目).value <> "" Then
            
        To勘定科目 = "収入" & "/" _
                        & rowSource.Range(enumSource.収入科目).value _
                        & "/" & rowSource.Range(enumSource.収入補助科目).value
    Else
        To勘定科目 = "?/?"
    End If

End Function


Public Function ToDate(ByVal YY As Long, ByVal MM As Long, ByVal DD As Long) As Date
    ' 年YYと月MMと日DDから日付を生成し西暦のDateとして返す
    ' 現金出納帳の年YYは和暦であるはずだから、適切に変換する。
    ' 令和を西暦年に変換するには（手抜きだとしりつつ）YYに整数2018を+することにした
    Dim d As Date
    d = DateSerial(YY + 2018, MM, DD)
    ToDate = d
End Function


Public Sub 入出金記録をソートする(ByVal wsCashTarget As Worksheet)
    Dim tblTarget As ListObject
    Set tblTarget = wsCashTarget.ListObjects(RecTableName)
    With tblTarget
        .Range.Sort key1:=.ListColumns(enumTarget.勘定科目), order1:=xlAscending, _
                    key2:=.ListColumns(enumTarget.年月日), order2:=xlAscending, _
                    Header:=xlYes
    End With
End Sub


'『現金出納記録』ワークシートに[テーブル勘定科目ごとの小計]がある。
'ここにデータを埋める。
'『現金出納記録』ワークシートに[テーブル現金出納記録]があるから、その行を
'勘定科目ごとの別シートに小分けして行をコピーし、テーブルに変換して集計を算出する。
'そして勘定科目と小計した金額を[テーブル現金出納記録]に書き込む。
'勘定科目ごとの別シートは集計が済めば無用なので削除する。

Public Sub 小計の表を作る(ByRef ws現金出納記録 As Worksheet)
    Dim tbl記録 As ListObject
    Set tbl記録 = ws現金出納記録.ListObjects(RecTableName)
    Dim tbl小計 As ListObject
    Set tbl小計 = ws現金出納記録.ListObjects(SumTableName)
    
    '[テーブル勘定科目ごとの小計]のデータ行をいったん消去する
    If tbl小計.ListRows.Count <> 0 Then
        tbl小計.DataBodyRange.Delete
    End If
    
    '勘定科目の一覧（重複なし）を取得する
    Dim unique勘定科目名の列 As Variant
    unique勘定科目名の列 = KzRange.KzGetUniqueItems(tbl記録.ListColumns(1).DataBodyRange)
    
    '勘定科目ごとの小計を算出して[テーブル勘定科目ごとの小計]に行として挿入する
    Dim column As Variant
    For Each column In unique勘定科目名の列
        Call 勘定科目の小計を算出する(ws現金出納記録, column)
    Next

End Sub


Public Sub Test小計の表を作る()
    Call BbLog.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(TheSheetName)
    Call 小計の表を作る(ws)
End Sub


'==============================================================================

Public Sub 勘定科目の小計を算出する(ByRef ws As Worksheet, ByVal 勘定科目名 As String)
    'Debug.Print "小計を算出する: " & 勘定科目名
    'Debug.Print "enumTarget.勘定科目 = " & enumTarget.勘定科目
    
    Dim tbl記録 As ListObject
    Set tbl記録 = ws.ListObjects(RecTableName)
    Dim tbl小計 As ListObject
    Set tbl小計 = ws.ListObjects(SumTableName)
    
    '勘定科目名をキーとするフィルタを[テーブル現金出納記録]に適用する。
    '指定された勘定科目に該当する行を選別しながら入出金データの行を走査する。
    '貸方金額の合計と借方金額の合計を算出する。
    '全部の行を調べ終わったら[テーブル勘定科目ごとの小計]に行をAddして
    '勘定科目と貸方金額の合計と借方金額の合計を書き出す。
    
    Dim debt As Long: debt = 0  '借方（左）
    Dim credit As Long: credit = 0  '貸方（右）
    Dim acc As String
    
    With tbl記録
        'フィルタを適用する
        .Range.AutoFilter 1, 勘定科目名
        'データ行ぜんぶを走査する
        Dim i As Long
        For i = 1 To .ListRows.Count
            'フィルタによって浮かび上がった行を選択する
            If .ListRows(i).Range.EntireRow.Hidden = False Then
                '小計を算出する
                acc = .ListRows(i).Range(enumTarget.勘定科目).value
                If acc Like "収入/*" Then
                    '勘定科目が収入なら
                    debt = debt + .ListRows(i).Range(enumTarget.借方金額) - .ListRows(i).Range(enumTarget.貸方金額)
                ElseIf acc Like "支出/*" Then
                    '勘定科目が支出なら
                    credit = credit + .ListRows(i).Range(enumTarget.貸方金額) - .ListRows(i).Range(enumTarget.借方金額)
                End If
                'Debug.Print acc & " " & debt & " " & credit
            End If
        Next
        'フィルタを解除
        .Range.AutoFilter enumTarget.勘定科目
    End With
    
    '[テーブル勘定科目ごとの小計]に行を挿入して算出した小計を書き込む
    With tbl小計.ListRows.Add
        .Range(enumTarget.勘定科目).value = acc
        .Range(enumTarget.借方金額).value = debt
        .Range(enumTarget.貸方金額).value = credit
    
        '金額のセルの書式を設定する（通貨、3桁毎カンマで区切る、\無し）
        .Range(enumTarget.借方金額).NumberFormatLocal = "#,###"
        .Range(enumTarget.貸方金額).NumberFormatLocal = "#,###"
    End With
End Sub

' 勘定科目の小計を算出するSubをテストする
Public Sub Test勘定科目の小計を算出する()
    Call BbLog.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TheSheetName)
    '[テーブル勘定科目ごとの小計]のデータ行をいったん消去する
    Dim tbl小計 As ListObject
    Set tbl小計 = ws.ListObjects(SumTableName)
    If tbl小計.ListRows.Count <> 0 Then
        tbl小計.DataBodyRange.Delete
    End If
    Call 勘定科目の小計を算出する(ws, "支出/慶弔費/慶弔費")
    Call 勘定科目の小計を算出する(ws, "支出/事務費/振込手数料")
    Call 勘定科目の小計を算出する(ws, "収入/会費/A会員")
End Sub


'==============================================================================


Public Sub 現金出納記録ワークシートが無ければ作る(ByRef Workbook As Workbook, ByVal sheetName As String)
    'ThisWorkbookのなかに【現金出納記録】ワークシートがすでにあるかどうかを調べる。
    'もしもすでにあったらなにもせずにおしまい。
    'まだ無かったら【現金出納記録】ワークシートを作り、"CashbookProj"のなかに
    '用意された【現金出納記録テンプレート】ワークシートの内容をコピーして保存する。
    If Not BbWorksheet.IsWorksheetPresentInWorkbook(Workbook, sheetName) Then
        Call BbLog.Info("AccountSum", "現金出納記録ワークシートが無ければ作る", sheetName & "ワークシートが未だ無いので作ります")
        Dim wsNew As Worksheet: Set wsNew = Worksheets.Add
        wsNew.Name = sheetName
        wsNew.Activate
        
        '[テーブル勘定科目ごとの小計]を作る
        Range("A2").value = "勘定科目"
        Range("B2").value = "列2"
        Range("C2").value = "列3"
        Range("D2").value = "列4"
        Range("E2").value = "列5"
        Range("F2").value = "列6"
        Range("G2").value = "列7"
        Range("H2").value = "列8"
        Range("I2").value = "借方金額"
        Range("J2").value = "貸方金額"
        Range("A3").value = "?"
        Range("I3").value = 0
        Range("J3").value = 0
        With wsNew
            .ListObjects.Add 1, Range("A3").CurrentRegion
            .ListObjects(1).Name = SumTableName
            .ListObjects(1).TableStyle = "TableStyleLight8"
            .ListObjects(1).ShowTotals = True
        End With
        Range("A1").value = "勘定科目ごとの小計"
        Range("A1").Style = "見出し 2"
        
        
        '[テーブル現金出納記録]を作る
        Range("A9").value = "勘定科目"
        Range("B9").value = "年月日"
        Range("C9").value = "収入科目"
        Range("D9").value = "収入補助科目"
        Range("E9").value = "支出科目"
        Range("F9").value = "支出補助科目"
        Range("G9").value = "収支報告単位"
        Range("H9").value = "摘要"
        Range("I9").value = "借方金額"
        Range("J9").value = "貸方金額"
        Range("A10").value = "?"
        Range("B10").value = ""
        Range("C10").value = ""
        Range("D10").value = ""
        Range("E10").value = ""
        Range("F10").value = ""
        Range("G10").value = ""
        Range("H10").value = ""
        Range("I10").value = 0
        Range("J10").value = 0
        With wsNew
            .ListObjects.Add 1, Range("A10").CurrentRegion
            .ListObjects(2).Name = RecTableName
            .ListObjects(2).TableStyle = "TableStyleLight9"
            .ListObjects(2).ShowTotals = True
        End With
        Range("A8").value = "現金出納記録"
        Range("A8").Style = "見出し 2"
        
        'ウインドウ枠を固定する　C3セルで
        Range("C3").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        
        
    End If
    
End Sub

'==============================================================================

'【現金出納記録】ワークシートの[テーブル勘定科目毎の小計]を検索して
'指定された勘定科目の小計金額を返す。
'例　勘定科目名「支出/会議費/役員会費」にたいして 33000 が返される
Public Function Get小計(ByVal 勘定科目名 As String) As Long
    Dim ws As Worksheet: Set ws = Worksheets(TheSheetName)
    Dim tbl小計 As ListObject: Set tbl小計 = ws.ListObjects(SumTableName)
    Dim val As Long: val = 0
    With tbl小計
        Dim i As Long
        For i = 1 To .ListRows.Count
            Dim acc As String: acc = .ListRows(i).Range(enumTarget.勘定科目).value
            If acc Like 勘定科目名 Then
                If acc Like "収入/*" Then
                    val = .ListRows(i).Range(enumTarget.借方金額)
                ElseIf acc Like "支出/*" Then
                    val = .ListRows(i).Range(enumTarget.貸方金額)
                Else
                    val = -1
                End If
                Exit For
            End If
        Next
    End With
    Get小計 = val
End Function


Public Sub Test_Get小計()
    Call BbLog.Clear
    Dim val As Long
    val = Get小計("支出/慶弔費/慶弔費")
    Debug.Assert val > 0
End Sub





