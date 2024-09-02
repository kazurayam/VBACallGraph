Attribute VB_Name = "BbRangeTest"
Option Explicit
Option Private Module

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

'@TestMethod("GetUniqueItems関数をテストする")
Private Sub Test_GetUniqueItems()
    Debug.Print String(300, vbCrLf)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("テストデータ")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("テーブル1")
    Dim item As Variant
    Dim uniqueItems As Variant
    uniqueItems = BbRange.GetUniqueItems(tbl.ListColumns(1).DataBodyRange)
    Dim i As Long: i = 0
    For Each item In uniqueItems
        i = i + 1
        Debug.Print i & " " & item
    Next

    Assert.AreEqual CLng(2), i
End Sub
