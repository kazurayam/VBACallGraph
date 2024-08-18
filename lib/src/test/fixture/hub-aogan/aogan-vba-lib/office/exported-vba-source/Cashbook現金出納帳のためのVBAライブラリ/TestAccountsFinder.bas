Attribute VB_Name = "TestAccountsFinder"
Option Explicit

Option Private Module

' TestAccountsFinder : AccountsFinder�I�u�W�F�N�g�����j�b�g�e�X�g����

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
    
    Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Set ws = wb.Worksheets("�����o�[��")
    Set tbl = ws.ListObjects("CashbookTable1")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'TearDown
    Application.DisplayAlerts = False ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肷��
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

'@TestMethod("�����Ȃ��FrpUnit=*, positiveLike=True�̏ꍇ")
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

'@TestMethod("��������FrpUnit=���k�u���b�N�u�K��, positiveLike=False�̏ꍇ")
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
    Set dic = af.FindAccounts("���k�u���b�N�u�K��", False)
    'Assert:
    Assert.AreEqual CLng(21), dic.Count
    Debug.Print dic.Count
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
