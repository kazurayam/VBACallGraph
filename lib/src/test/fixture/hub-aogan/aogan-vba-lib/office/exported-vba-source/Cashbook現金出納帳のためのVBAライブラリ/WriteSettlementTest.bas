Attribute VB_Name = "WriteSettlementTest"
Option Explicit
Option Private Module

' TestWriteSettlement : AccoutSumモジュールをテストする

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'==============================================================================

'@TestMethod("【現金出納記録】ワークシートを作る。Get小計関数をテストする。")
Private Sub Test現金出納記録を作る()
    On Error GoTo TestFail
    '令和5年度予算（案）・令和4年度決算（案）.xlsmワークブックに作りこまれた
    '決算書作成モジュールのコードと同じものをRubberduckの制御下で実行する。

    Call BbLog.Clear      ' イミディエイトウィンドウを初期化する
    Call BbLog.Info("TestWriteSettlement", "Test現金出納記録を作る", "現金出納記録ワークシートを更新します")
    
    'このワークブックのなかに【現金出納記録】ワークシートを作る
    Call WriteSettlement.現金出納記録ワークシートが無ければ作る(ThisWorkbook, "現金出納記録")
    Dim ws現金出納記録 As Worksheet
    Set ws現金出納記録 = ThisWorkbook.Worksheets("現金出納記録")
    
    '現金出納帳ワークブックを開く
    Dim sourcePath As String
    sourcePath = BbUtil.ResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2")
    Call BbLog.Info("TestWriteSettlement", "Test現金出納記録を作る", "入力: " & sourcePath)
    
    Dim wb現金出納帳 As Workbook: Set wb現金出納帳 = Workbooks.Open(sourcePath)
    Dim wsSource As Worksheet: Set wsSource = wb現金出納帳.Worksheets("現金出納帳")
    
    '現金出納帳のデータを選別しながらimportして[テーブル現金出納記録]に転記する
    '令和４年４月１日から令和５年３月３１日までの入出金を選択する
    '収支報告単位が「東北ブロック講習会」であるレコードを除外する。
    Call WriteSettlement.入出金記録を取り込む(wsSource, ws現金出納記録, _
                                periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#, _
                                ofReportingUnit:="東北ブロック講習会", positiveLike:=False)
    
    '[現金出納記録]テーブルのデータ行を並べ替える、勘定科目毎の明細を読み取れるように
    Call WriteSettlement.入出金記録をソートする(ws現金出納記録)
    
    '[テーブル勘定科目ごとの小計]を更新する
    Call WriteSettlement.小計の表を作る(ws現金出納記録)
    
    ' 「変更内容を保存しますか」ダイアログを表示しないように設定して
    Application.DisplayAlerts = False
    '現金出納帳ワークブックを閉じる
    wb現金出納帳.Close
    Set wb現金出納帳 = Nothing
    
    Call BbLog.Info("TestWriteSettlement", "Test現金出納記録を作る", "現金出納記録ワークシートを更新しました")


    'Get小計関数をテストする
    Dim val As Long
    val = WriteSettlement.Get小計("支出/慶弔費/慶弔費")
    Assert.IsTrue val > 0, "Get小計()がゼロを返した"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
