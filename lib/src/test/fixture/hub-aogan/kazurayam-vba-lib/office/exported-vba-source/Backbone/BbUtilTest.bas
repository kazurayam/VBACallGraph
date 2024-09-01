Attribute VB_Name = "BbUtilTest"
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



'@TestMethod("Function VarTypeAsStringをユニットテストする")
Private Sub Test_VarTypeAsString()
    On Error GoTo TestFail
    'Arrange:
    Dim integerVar As Integer: integerVar = 0
    Dim longVar As Long: longVar = 0
    Dim doubleVar As Double: doubleVar = 0
    Dim stringVar As String: stringVar = ""
    Dim booleanVar As Boolean: booleanVar = False
    Dim dateVar As Date: dateVar = Date
    Dim objectVar As Object: Set objectVar = ThisWorkbook
    ' Dim variantVar As Variant: variantVar = "123"
    Dim h() As String
    Dim i() As Integer
    'Act:
    'Assert:
    Assert.AreEqual "Integer", VarTypeAsString(integerVar)
    Assert.AreEqual "Long", VarTypeAsString(longVar)
    Assert.AreEqual "Double", VarTypeAsString(doubleVar)
    Assert.AreEqual "String", VarTypeAsString(stringVar)
    Assert.AreEqual "Boolean", VarTypeAsString(booleanVar)
    Assert.AreEqual "Date", VarTypeAsString(dateVar)
    Assert.AreEqual "Object", VarTypeAsString(objectVar)
    ' Assert.AreEqual "Variant", VarTypeAsString(variantVar)
    Assert.AreEqual "String()", VarTypeAsString(h)
    Assert.AreEqual "Integer()", VarTypeAsString(i)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
