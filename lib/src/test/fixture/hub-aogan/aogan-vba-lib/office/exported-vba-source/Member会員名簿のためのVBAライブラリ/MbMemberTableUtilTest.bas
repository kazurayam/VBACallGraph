Attribute VB_Name = "MbMemberTableUtilTest"
Option Explicit
Option Private Module

' MemberUtils���e�X�g����
' Rubberduck�����s�\��unittest�Ƃ��Ď������Ă���B

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

'@TestMethod("FetchMemberTable�֐����e�X�g����")
Private Sub Test_FetchMemberTable()
    Call BbLog.Clear
    'Arrange:
    '�u�O���t�@�C���̃p�X�v�V�[�g��B2�Z���ɉ������Excel�t�@�C���̃p�X�������Ă���͂��A�����ǂݎ��
    Dim filePath As String: filePath = BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B2")
    'Act:
    Dim tbl As ListObject: Set tbl = MbMemberTableUtil.FetchMemberTable(filePath, "R6�N�x", ThisWorkbook)
    'Assert:
    Debug.Print ("����e�[�u���̍s��: " & tbl.ListRows.Count)
    Assert.IsTrue tbl.ListRows.Count > 0
    
End Sub
