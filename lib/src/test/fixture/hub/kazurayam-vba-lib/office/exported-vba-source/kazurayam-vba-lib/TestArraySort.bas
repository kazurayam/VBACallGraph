Attribute VB_Name = "TestArraySort"

Option Explicit
Option Private Module

' TestArraySort: ArraySortモジュールをテストする
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


'@TestMethod("配列をソートする")
Private Sub Test_InsertionSort()
    On Error GoTo TestFail
    'Arrange:
    KzCls
    Dim data() As Variant
    data = Array("ddd", "xxx", "jjj", "aaa", "9")
    ' 並べ替える
    Call InsertionSort(data, LBound(data), UBound(data))
    Dim d As Variant
    For Each d In data
        Debug.Print d
        ' 0 1 2 3 4 5 6 7 8 9
    Next
    'Assert:
    'Act:
    'Assert:
    Assert.AreEqual "9", data(0)
    Assert.AreEqual "aaa", data(1)
    Assert.AreEqual "ddd", data(2)
    Assert.AreEqual "jjj", data(3)
    Assert.AreEqual "xxx", data(4)
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



