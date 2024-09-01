Attribute VB_Name = "CbHelperTest"
Option Explicit
Option Private Module

' CbHelperTest : CbHelperモジュールが定義するSubやFunctionをユニットテストする

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Dim wb As Workbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, "現金出納帳ファイルのパス", "B2"))
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
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("Sub PrintAccountsをテストする")
Private Sub Test_PrintAccounts()
    On Error GoTo TestFail
    Call BbLog.Clear
    'Arrange:
    Dim cb As Cashbook: Set cb = CbFactories.CreateCashbook(wb, "現金出納帳", "CashbookTable1")
    Dim rpUnit As String: rpUnit = "東北ブロック講習会"
    Dim positiveLike As Boolean: positiveLike = False
    'Act:
    Call PrintAccounts(cb, rpUnit, positiveLike)
    'Assert:
    Assert.IsTrue True   ' placeholder
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


