Attribute VB_Name = "MbMemberTableUtilTest"
Option Explicit
Option Private Module

' MemberUtilsをテストする
' Rubberduckが実行可能なunittestとして実装してある。

'@TestModule
'@Foler("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub MoudleInitialize()
    'this method runs once per module
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'This method runs after every test in the module
End Sub

'@TestMethod("FetchMemberTable関数をテストする")
Private Sub Test_FetchMemberTable()
    Call BbLog.Clear
    'Arrange:
    '「外部ファイルのパス」シートのB2セルに会員名簿Excelファイルのパスが書いてあるはず、それを読み取る
    Dim filePath As String: filePath = BbUtil.ResolveExternalFilePath(ThisWorkbook, "外部ファイルのパス", "B2")
    'Act:
    Dim tbl As ListObject: Set tbl = MbMemberTableUtil.FetchMemberTable(filePath, "R6年度", ThisWorkbook)
    'Assert:
    Debug.Print ("会員テーブルの行数: " & tbl.ListRows.Count)
    Assert.IsTrue tbl.ListRows.Count > 0
    
End Sub
