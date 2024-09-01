Attribute VB_Name = "CashbookTest"
Option Explicit
Option Private Module

' TestCashbook : Cashbook�N���X�����j�b�g�e�X�g����

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Dim wb As Workbook
Dim ws As Worksheet
Dim tbl As ListObject
Dim cb As Cashbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")

    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
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


'@TestMethod("�����o�[�����[�N�V�[�g�̃e�[�u��1�ɃA�N�Z�X����e�X�g")
Private Sub Test_AccessingTable1()
    On Error GoTo TestFail
    'Arrange:
    'Act
    'Assert:
    Assert.IsTrue tbl.ListRows.Count > 0, "�e�[�u��1������ۂ�"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Cashbook�C���X�^���X�𐶐���Count���\�b�h��GetCash���\�b�h���e�X�g����")
Private Sub Test_Initialize_Cout_GetCash()
    On Error GoTo TestFail
    'Arrange:
    'Act:
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    'Assert:
    Assert.IsTrue cb.Count > 0, "Cashbook�I�u�W�F�N�g���������Count���[����"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Cashbook�I�u�W�F�N�g�𐶐����āACount�v���p�e�B��ǂ݂�����Print����")
Private Sub TasteCashbook()
    'Act:
    Dim cb As Cashbook: Set cb = Factories.CreateCashbook(wb, "�����o�[��", "CashbookTable1")
    'Assert:
    Debug.Print "cb.Count=" & cb.Count
    'TearDown
End Sub

