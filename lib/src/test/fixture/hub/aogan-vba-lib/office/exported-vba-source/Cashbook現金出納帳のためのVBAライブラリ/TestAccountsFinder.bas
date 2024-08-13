Attribute VB_Name = "TestAccountsFinder"
Option Explicit

Option Private Module

' TestAccountsFinder : AccountsFinderオブジェクトをユニットテストする

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

'@TestMethod("引数なし：rpUnit=*, positiveLike=Trueの場合")
Private Sub Test_FindAccounts()
    On Error GoTo TestFail
    'Arrange:
    Call KzCls
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act:
    Dim af As AccountsFinder: Set af = New AccountsFinder
    Call af.Initialize(cb)
    Dim dic As Dictionary
    Set dic = af.FindAccounts()
    'Assert:
    Assert.AreEqual CLng(26), dic.Count
    Debug.Print dic.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("引数あり：rpUnit=東北ブロック講習会, positiveLike=Falseの場合")
Private Sub Test_FindAccounts_NotLike()
    On Error GoTo TestFail
    'Arrange:
    Call KzCls
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Act:
    Dim af As AccountsFinder: Set af = New AccountsFinder
    Call af.Initialize(cb)
    Dim dic As Dictionary
    Set dic = af.FindAccounts("東北ブロック講習会", False)
    'Assert:
    Assert.AreEqual CLng(21), dic.Count
    Debug.Print dic.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
