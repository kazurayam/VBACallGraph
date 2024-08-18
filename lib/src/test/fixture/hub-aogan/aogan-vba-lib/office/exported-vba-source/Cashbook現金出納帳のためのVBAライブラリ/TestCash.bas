Attribute VB_Name = "TestCash"
Option Explicit
Option Private Module

' TestCash : Cashクラスをユニットテストする

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Dim wb As Workbook
Dim ws As Worksheet
Dim tbl As ListObject

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Set ws = wb.Worksheets("現金出納帳")
    Set tbl = ws.ListObjects("CashbookTable1")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
    Set wb = Nothing
    Set ws = Nothing
    Set tbl = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("現金出納帳ワークシートのテーブル1にアクセスするテスト")
Private Sub TestAccessTable1()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    'Assert:
    Assert.isTrue tbl.ListRows.Count > 0, "テーブルが空っぽかもしれない"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




'@TestMethod("Cashオブジェクトを作りプロパティを読みだすテスト　支出の場合")
Private Sub TestExpense()
    On Error GoTo TestFail
    'Arrange:
    Dim data As Range: Set data = tbl.ListRows(10).Range '支出の一例
    Dim cs As Cash: Set cs = New Cash
    'Act:
    Call cs.Initialize(data)
    'Assert:
        'Debug.Print "cs.ReceiptNo is " + KzVarTypeAsString(cs.ReceiptNo)
        'Debug.Print "5 is " + KzVarTypeAsString(5)
        'Debug.Print "CLng(5) is " + KzVarTypeAsString(CLng(5))
    Assert.AreEqual CLng(5), cs.ReceiptNo     '領収書No.
    Assert.AreEqual CLng(4), cs.YY            '年度（令和）
    Assert.AreEqual CLng(5), cs.MM            '月
    Assert.AreEqual CLng(24), cs.DD           '日
    Assert.AreEqual "", cs.incomeAccount            '収入科目
    Assert.AreEqual "", cs.IncomeSubAccount         '収入補助科目
    Assert.AreEqual "事業費", cs.expenseAccount     '支出科目
    Assert.AreEqual "広報費", cs.ExpenseSubAccount  '支出補助科目
    Assert.AreEqual "会報", cs.ReportingUnit        '収支報告単位
    Assert.AreEqual "会報第73号印刷代　凸版メディア", cs.Description   '適用
    Assert.AreEqual CLng(0), cs.IncomeAmount              '借方金額
    Assert.AreEqual CLng(495000), cs.ExpenseAmount        '貸方金額
    Assert.AreEqual "支出/事業費/広報費", cs.itsAccount.ToString
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Cashオブジェクトを作りプロパティを読みだすテスト　収入の場合")
Private Sub TestIncome()
    On Error GoTo TestFail
    'Arrange:
    Dim data As Range: Set data = tbl.ListRows(12).Range '収入の一例
    '    4   5   28  雑収入  セミナー参加料          眼科フォーラム  ｸﾏｶﾞｲｼｭﾝｲﾁ　セミナー受講料   2,000
    Dim cs As Cash: Set cs = New Cash
    'Act:
    Call cs.Initialize(data)
    'Assert:
    Assert.AreEqual CLng(0), cs.ReceiptNo     '領収書No.
    Assert.AreEqual CLng(4), cs.YY            '年度（令和）
    Assert.AreEqual CLng(5), cs.MM            '月
    Assert.AreEqual CLng(28), cs.DD           '日
    Assert.AreEqual "雑収入", cs.incomeAccount            '収入科目
    Assert.AreEqual "セミナー参加料", cs.IncomeSubAccount         '収入補助科目
    Assert.AreEqual "", cs.expenseAccount     '支出科目
    Assert.AreEqual "", cs.ExpenseSubAccount  '支出補助科目
    Assert.AreEqual "眼科フォーラム", cs.ReportingUnit        '収支報告単位
    Assert.AreEqual "ｸﾏｶﾞｲｼｭﾝｲﾁ　セミナー受講料", cs.Description   '適用
    Assert.AreEqual CLng(2000), cs.IncomeAmount            '借方金額
    Assert.AreEqual CLng(0), cs.ExpenseAmount        '貸方金額
    
    Assert.AreEqual "収入/雑収入/セミナー参加料", cs.itsAccount.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashクラスのColumnHeaderメソッドとToStringメソッドをテストする")
Private Sub TestToString()
    On Error GoTo TestFail
    'Arrange:
    Call KzCls
    Dim data As Range: Set data = tbl.ListRows(12).Range '収入の一例
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    'Act:
    Dim ch As String: ch = cs.ColumnHeader()
    Dim s As String: s = cs.ToString()
    Debug.Print ch
    Debug.Print s
    'Assert:
    Assert.isTrue InStr(ch, "領収書") <> 0, "ColumnHeader"
    Assert.isTrue InStr(s, " ") <> 0, "領収書No."
    Assert.isTrue InStr(s, "4") <> 0, "年"
    Assert.isTrue InStr(s, "5") <> 0, "月"
    Assert.isTrue InStr(s, "28") <> 0, "日"
    Assert.isTrue InStr(s, "雑収入") <> 0, "収入科目"
    Assert.isTrue InStr(s, "セミナー参加料") <> 0, "収入補助科目"
    Assert.isTrue InStr(s, " ") <> 0, "支出科目"
    Assert.isTrue InStr(s, " ") <> 0, "支出補助科目"
    Assert.isTrue InStr(s, "眼科フォーラム") <> 0, "収支報告単位"
    Assert.isTrue InStr(s, "ｸﾏｶﾞｲｼｭﾝｲﾁ　セミナー受講料") <> 0, "適用"
    Assert.isTrue InStr(s, "2000") <> 0, "借方金額"
    Assert.isTrue InStr(s, " ") <> 0, "貸方金額"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashクラスのToDateメソッドをテストする")
Private Sub TestToDate()
    On Error GoTo TestFail
    'Arrange:
    Dim data As Range: Set data = tbl.ListRows(12).Range '収入の一例
    Dim cs As Cash: Set cs = New Cash
    Call cs.Initialize(data)
    'Act:
    Debug.Print cs.ToDate()
    'Assert:
    Assert.AreEqual #5/28/2022#, cs.ToDate()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
