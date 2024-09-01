Attribute VB_Name = "CashbookTest"
Option Explicit
Option Private Module

' TestCashbook : Cashbookクラスをユニットテストする

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Dim wb As Workbook
Dim ws As Worksheet
Dim tbl As ListObject
Dim cb As Cashbook

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


'@TestMethod("現金出納帳ワークシートのテーブル1にアクセスするテスト")
Private Sub Test_AccessingTable1()
    On Error GoTo TestFail
    'Arrange:
    'Act
    'Assert:
    Assert.IsTrue tbl.ListRows.Count > 0, "テーブル1が空っぽだ"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Cashbookインスタンスを生成しCountメソッドとGetCashメソッドをテストする")
Private Sub Test_Initialize_Cout_GetCash()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Assert:
    Assert.IsTrue cb.Count > 0, "Cashbookオブジェクトを作ったがCountがゼロだ"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Cashbookオブジェクトを生成して、Countプロパティを読みだしてPrintする")
Private Sub TasteCashbook()
    'Act:
    Dim cb As Cashbook: Set cb = Factories.CreateCashbook(wb, "現金出納帳", "CashbookTable1")
    'Assert:
    Debug.Print "cb.Count=" & cb.Count
    'TearDown
End Sub

