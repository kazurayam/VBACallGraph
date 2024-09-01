Attribute VB_Name = "CashListTest"
Option Explicit
Option Private Module

' TestCashList: CashListクラスをユニットテストする

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Dim wb As Workbook
Dim ws As Worksheet
Dim tbl As ListObject
Dim cb As Cashbook
Dim cs As CashSelector
    
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
    Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Set cs = New CashSelector
    Call cs.Initialize(cb)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set cb = Nothing
    Set cs = Nothing
End Sub


'@TestMethod("CashListオブジェクトのCountプロパティをテストする")
Private Sub TestCount()
    On Error GoTo TestFail
    'Arrange:
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "雑収入", "セミナー参加料", "眼科フォーラム")
    'Act:
    'Assert:
    Assert.AreEqual CLng(4), selected.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashListオブジェクトのItems関数をテストする")
Private Sub TestItems()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "雑収入", "セミナー参加料", "眼科フォーラム")
    'Act:
    'Assert:
    Assert.AreEqual CLng(4), selected.Count
    Dim i As Long
    For i = 1 To selected.Count
        Debug.Print selected.Items(i).ToString()
    Next i
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CashListオブジェクトのSumOfIncomeAmountメソッドをテストする")
Private Sub TestSumOfIncomeAmount()
    On Error GoTo TestFail
    'Arrange:
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Income, "雑収入", "セミナー参加料", "眼科フォーラム")
    Call BbLog.Clear
    Debug.Print selected.ToString()
    'Act:
    Dim sum As Long
    sum = selected.SumOfIncomeAmount()
    'Assert:
    Assert.AreEqual CLng(56000), sum
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashListオブジェクトのSumOfExpenseAmountメソッドをテストする")
Private Sub TestSumOfExpenseAmount()
    On Error GoTo TestFail
    'Arrange:
    Dim selected As CashList
    Set selected = cs.SelectCashList(AccountType.Expense, "事業費", "公衆衛生費")
    Call BbLog.Clear
    Debug.Print selected.ToString()
    'Act:
    Dim sum As Long
    sum = selected.SumOfExpenseAmount()
    'Assert:
    Assert.AreEqual CLng(2), selected.Count
    Debug.Print sum
    Assert.AreEqual CLng(540000), sum
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

