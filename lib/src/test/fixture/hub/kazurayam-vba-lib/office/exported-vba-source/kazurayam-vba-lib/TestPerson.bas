Attribute VB_Name = "TestPerson"
Option Explicit

Option Private Module

' TestPerson: Personクラスをテストする

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


'@TestMethod("Sub Initializeをテストする")
Private Sub Test_Initialize_Alice()
    On Error GoTo TestFail
    KzCls
    'Arrange:
    Dim alice As Person
    Set alice = New Person
    'Act:
    Call alice.Initialize("Alice", 16)
    'Assert:
    Debug.Print "Name: " & alice.Name
    Debug.Print "Age: " & alice.GetAge()
    Assert.IsTrue alice.Name Like "Alice"
    Assert.AreEqual alice.GetAge(), CLng(16)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


