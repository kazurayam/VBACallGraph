Attribute VB_Name = "CashSelectorTest"
Option Explicit
Option Private Module

' TestCashSelector : CashSelectorクラスをユニットテストする

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
    
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Set ws = wb.Worksheets("現金出納帳")
    Set tbl = ws.ListObjects("CashbookTable1")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'TearDown
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

'@TestMethod("CashSelectorインスタンスを生成しSelectCashListメソッドをテストする - Incomeの場合")
Private Sub Test_SelectCashList_ofIncome()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act:
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "雑収入", "セミナー参加料", "眼科フォーラム")
    'Assert:
    Assert.AreEqual CLng(4), selected.Count
    Debug.Print selected.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CashSelectorインスタンスを生成しSelectCashListメソッドをテストする - Not Likeの場合")
Private Sub Test_SelectCachList_NotLike()
    ' positiveLikeパラメータにfalseを指定して「東北ブロック講習会」を除外した事務費／通信費をリストする
    On Error GoTo TestFail:
    'Arrange
    Call BbLog.Clear
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "事務費", "通信費", "東北ブロック講習会", False)
    'Assert
    Debug.Print selected.ToString()
    Assert.IsTrue (CLng(0) < selected.Count)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CashSelectorインスタンスを生成しSelectCashListByMatchingDescriptionメソッドをテストする")
Private Sub Test_SelectCashListByMatchingDescription()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act:
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    Set selected = cs.SelectCashListByMatchingDescription(AccountType.Income, "会費", "A会員", "ﾀﾑﾗ ﾏｻﾄ")
    'Assert:
    Assert.AreEqual CLng(1), selected.Count
    Debug.Print selected.ToString()
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("CashSelectorオブジェクトのSelectCashListメソッドを味見する:Incomeの場合")
Private Sub TasteCashSelector_SelectCashList_ofIncome()
    'Arrange
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Dim selected As CashList
    'Act
    Set selected = cs.SelectCashList(AccountType.Income, "雑収入", "セミナー参加料", "眼科フォーラム")
    Debug.Print selected.Count
    Debug.Print selected.ToString()
End Sub


'@TestMethod("CashSelectorオブジェクトのSelectCashListOfIncomeメソッドを味見する")
Private Sub TasteCashSelector_SelectCashList_ofExpense()
    'Arrange
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    'Act
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "事業費", "広報費", "会報")
    'Assert
    Debug.Print selected.Count
    Debug.Print selected.ToString()
End Sub



