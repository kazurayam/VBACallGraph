Attribute VB_Name = "TestDicSort"
Option Explicit
Option Private Module

' TestDicSort: 連想配列Dictionaryをキーの昇順に並べかえるSub DicSortをテストする

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


'@TestMethod("keyとvalueがどちらもString型であるDictionaryをソートする")
Private Sub TestDictionaryOfStringString()
    On Error GoTo TestFail
    'Arrange:
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    dic("g") = "gggg"
    dic("9") = "999"
    dic("を") = "をををを"
    dic("4") = "444"
    dic("あ") = "ああああ"
    dic("(") = "(((("
    dic("a") = "aaaa"
    'Assert:
    Dim keys() As Variant    ' Keysという名前のdynamic arrayを宣言している
    keys = dic.keys
    Assert.AreEqual "g", keys(0)
    Assert.AreEqual "9", keys(1)
    Assert.AreEqual "を", keys(2)
    
    'Act:
    Dim output As String
    Dim key As Variant
    output = "##before------------" & vbNewLine
    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key
    
    ' now sort it
    Call DicSort(dic)
    
    output = output & vbNewLine & "##after-------------" & vbNewLine
    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key
    
    Call KzCls
    Debug.Print output
    
    'Assert:
    keys = dic.keys
    Assert.AreEqual "(", keys(0)
    Assert.AreEqual "4", keys(1)
    Assert.AreEqual "9", keys(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
