Attribute VB_Name = "決算書作成"
Option Explicit

' 決算書作成 Module
'
'青森県眼科医会の現金出納帳ワークブックに記録された入出金データを主な入力情報として
'『決算書』シートに数字を埋めることを最終目的とする。
'
'『実行環境』シートのなかに外部のワークブックの具体的なパスが書いてあるのを頼りとして
'現金出納帳ワークブックをopenする。
'
'現金出納帳ワークブックのなかから次の条件に該当する行を選択する。
'  1. 令和5年4月1日から令和6年3月31日までの期間に属する入出金　かつ
'  2. 【収支報告単位】が「東北ブロック講習会」ではない入出金
'選択した行を『入出金記録』シートのテーブルにコピーする。
'テーブルの行を勘定科目の昇順　＞　年月日の昇順で並べ替える。
'並べ替えると『入出金記録』が勘定科目ごとの明細情報として読むことができるようになる。

'『入出金記録』テーブルのデータから勘定科目毎の小計を算出し[勘定科目ごとの小計]テーブルに出力する。
'『決算書』ワークシートのセルに参照式を設定し、Functionを呼び出する。
'Functionは[小計テーブル]を検索して勘定科目ごとの小計金額を返す。


'==============================================================================
Public Sub 決算書作成()

    Call KzUtil.KzCls      ' イミディエイトウィンドウを初期化する
    Call KzUtil.KzLog("決算書作成", "決算書作成", "現金出納記録ワークシートを更新します")
    
    'このワークブックのなかに【現金出納記録】ワークシートを作る
    Call CashbookPrj.WriteSettlement.現金出納記録ワークシートが無ければ作る(ThisWorkbook, "現金出納記録")
    
    Dim ws現金出納記録 As Worksheet
    Set ws現金出納記録 = ThisWorkbook.Worksheets("現金出納記録")
    
    '現金出納帳ワークブックを開く
    Dim sourcePath As String
    sourcePath = KzUtil.KzResolveExternalFilePath(ThisWorkbook, "実行環境", "B2")
    Call KzUtil.KzLog("決算書作成", "決算書作成", "入力: " & sourcePath)
    
    Dim wb現金出納帳 As Workbook: Set wb現金出納帳 = Workbooks.Open(sourcePath)
    Dim wsSource As Worksheet: Set wsSource = wb現金出納帳.Worksheets("現金出納帳")
    
    '現金出納帳のデータを選別しながらimportして[テーブル現金出納記録]に転記する
    '令和5年４月１日から令和6年３月３１日までの入出金を選択する
    '収支報告単位が「東北ブロック講習会」であるレコードを除外する。
    Call CashbookPrj.WriteSettlement.入出金記録を取り込む(wsSource, ws現金出納記録, _
                                periodStart:=#4/1/2023#, periodEnd:=#3/31/2024#, _
                                ofReportingUnit:="東北ブロック講習会", positiveLike:=False)
    
    '[現金出納記録]テーブルのデータ行を並べ替える、勘定科目毎の明細を読み取れるように
    Call CashbookPrj.WriteSettlement.入出金記録をソートする(ws現金出納記録)
    
    '[テーブル勘定科目ごとの小計]を更新する
    Call CashbookPrj.WriteSettlement.小計の表を作る(ws現金出納記録)
    
    ' 「変更内容を保存しますか」ダイアログを表示しないように設定して
    Application.DisplayAlerts = False
    '現金出納帳ワークブックを閉じる
    wb現金出納帳.Close
    Set wb現金出納帳 = Nothing
    
    Call KzUtil.KzLog("決算書作成", "決算書作成", "現金出納記録ワークシートを更新しました")
    
End Sub

