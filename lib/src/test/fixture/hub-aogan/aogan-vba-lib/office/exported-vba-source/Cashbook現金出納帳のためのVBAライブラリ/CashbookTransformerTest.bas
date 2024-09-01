Attribute VB_Name = "CashbookTransformerTest"
Option Explicit

Option Private Module

' CashbookTransformerTest : CashbookTransformerクラスをユニットテストする

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
    
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, _
                                            "現金出納帳ファイルのパス", "B2"))
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
Private Sub Test_ByAccounts()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cashbook_ As Cashbook
    Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act:
    Dim cbTransformer As CashbookTransformer
    Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts()
    
    'Assert:
    Assert.IsTrue (CLng(0) < cashListDic.Count)
    Debug.Print cashListDic.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("引数あり：rpUnit=東北ブロック講習会, positiveLike=Falseの場合")
Private Sub Test_byAccounts_positiveLike_False()
    On Error GoTo TestFail
    'Arrange:
    Call BbLog.Clear
    Dim cashbook_ As Cashbook
    Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act:
    Dim cbTransformer As CashbookTransformer
    Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts("東北ブロック講習会", False)
    'Assert:
    Assert.IsTrue (CLng(0) < cashListDic.Count)
    Debug.Print cashListDic.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CashbookTransformerクラスを味見する")
Private Sub Taste_byAccounts_showCount()
    On Error GoTo TestFail
    'Arrange
    Dim cashbook_ As Cashbook: Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act
    Dim cbTransformer As CashbookTransformer: Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary: Set cashListDic = cbTransformer.ByAccounts("*", True)
    
    Dim accounts As Variant: Let accounts = cashListDic.Keys
    'キーを昇順にソートする
    Call BbArraySort.InsertionSort(accounts, LBound(accounts), UBound(accounts))
    
    Call BbLog.Clear
    Dim i As Long
    For i = LBound(accounts) To UBound(accounts)
        Dim account_ As String: account_ = accounts(i)
        Dim msg As String
        msg = account_ & ":" & CStr(cashListDic(account_).Count())
        Call BbLog.Info("TestCashbookTransformer", "Taste_byAccounts_showCount", msg)
    Next i
'0//:1
'収入/会費/A会員:44
'収入/会費/B会員:32
'収入/会費/C会員:1
'収入/補助金/補助金:1
'収入/雑収入/セミナー参加料:4
'収入/雑収入/広告料収入:12
'収入/雑収入/雑収入:1
'支出/事務費/消耗品費:27
'支出/事務費/通信費:58
'支出/事業費/公衆衛生費:1
'支出/事業費/学術費:17
'支出/事業費/広報費:3
'支出/事業費/庶務費:1
'支出/事業費/眼科コ・メディカル関係費:1
'支出/事業費/通信費:1
'支出/会議費/役員会費:1
'支出/出張費/出張旅費補助:1
'支出/慶弔費/慶弔費:1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FindKeysAs Stringをテストする")
Private Sub TestFindKeysAsString()
    On Error GoTo TestFail
    'Arrange
    Dim cashbook_ As Cashbook: Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act
    Dim cbTransformer As CashbookTransformer: Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)

    Call BbLog.Info("TestCashbookTransformer", "Test_FindKeysAsString", cbTransformer.FindKeysAsString())
'0//
'収入/会費/A会員
'収入/会費/B会員
'収入/会費/C会員
'収入/補助金/補助金
'収入/雑収入/セミナー参加料
'収入/雑収入/広告料収入
'収入/雑収入/雑収入
'支出/事務費/消耗品費
'支出/事務費/通信費
'支出/事業費/公衆衛生費
'支出/事業費/学術費
'支出/事業費/広報費
'支出/事業費/庶務費
'支出/事業費/眼科コ・メディカル関係費
'支出/事業費/通信費
'支出/会議費/役員会費
'支出/出張費/出張旅費補助
'支出/慶弔費/慶弔費
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AccountsFinderクラスを味見する ただし収支報告単位を特定して")
Private Sub Test_byAccounts_ReportingUnit()
    'Arrange
    Dim cashbook_ As Cashbook: Set cashbook_ = New Cashbook
    Call cashbook_.Initialize(tbl)
    'Act
    Dim cbTransformer As CashbookTransformer: Set cbTransformer = New CashbookTransformer
    Call cbTransformer.Initialize(cashbook_)
    
    Dim cashListDic As Dictionary
    Set cashListDic = cbTransformer.ByAccounts(ofReportingUnit:="東北ブロック講習会")
    'Assert
    Call BbLog.Clear
    Dim accounts As Variant
    accounts = cashListDic.Keys
    Dim i As Long
    For i = LBound(accounts) To UBound(accounts)
        Dim account As String: account = accounts(i)
        Dim msg As String: msg = i & " " & account & " " & cashListDic(account).Count()
        Call BbLog.Info("TestCashbookTransformer", "Test_byAccounts_ReportingUnit", msg)
    Next i
    
'支出/事務費/通信費:17
'支出/事業費/学術費:9
'支出/事業費/通信費:1

End Sub

