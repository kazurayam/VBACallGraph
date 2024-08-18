Attribute VB_Name = "TestFactories"
Option Explicit
Option Private Module

'TestFactories: Factoriesモジュールをユニットテストする

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private wb As Workbook
Private ws As Worksheet
Private tbl As ListObject
Private cb As Cashbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    '
    Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
    Set ws = wb.Worksheets("現金出納帳")
    Set tbl = ws.ListObjects("CashbookTable1")
    Set cb = New Cashbook
    Call cb.Initialize(tbl)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'Teardown
    Application.DisplayAlerts = False ' 「変更内容を保存しますか」ダイアログを表示しないように設定する
    wb.Close
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("CreateCashbook関数をテストする")
Private Sub TestCreateCashbook()
    'Act:
    Dim cbx As Cashbook: Set cbx = CreateCashbook(wb, "現金出納帳", "CashbookTable1")
    'Assert:
    Assert.AreEqual CLng(313), cbx.Count
End Sub


'@TestMethod("CreateAccountsFinder関数をテストする")
Private Sub TestCreateAccountsFinder()
    'Act
    Dim accFinder As AccountsFinder: Set accFinder = CreateAccountsFinder(cb)
    Call KzCls
    Debug.Print accFinder.FindKeysAsString
    '
End Sub


'@TestMethod("CreateCashSelector関数をテストする")
Private Sub TestCreateCashSelector()
    Dim cs As CashSelector: Set cs = CreateCashSelector(cb)
    Call KzCls
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "雑収入", "セミナー参加料", "眼科フォーラム")
    'Assert:
    Debug.Print selected.Count
    Assert.AreEqual CLng(4), selected.Count
End Sub

'@TestMethod("CreateEmptyCashList関数をテストする")
Private Sub TestCreateEmptyCashList()
    Call KzCls
    Dim cl As CashList
    Set cl = CreateEmptyCashList()
    Assert.AreEqual CLng(0), cl.Count
End Sub
