Attribute VB_Name = "TestKzWorksheet"
Option Explicit
Option Private Module

'Kz���W���[���ɏ����ꂽPublic��Sub��Function��Rubberduck���g���ă��j�b�g�e�X�g����

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


'@TestMethod("KzIsWorksheetPresentInWorkbook���e�X�g����")
Private Sub Test_KzIsWorksheetPresentInWorkbook()
    'Assert:
    Assert.IsTrue KzIsWorksheetPresentInWorkbook(ThisWorkbook, "�e�X�g�f�[�^")
    Assert.IsFalse KzIsWorksheetPresentInWorkbook(ThisWorkbook, "No Such Worksheet")
End Sub

'@TestMethod("KzCreateWorksheetInWorkbook���e�X�g����")
Private Sub Test_KzCreateWorksheetInWorkbook()
    'Arrange
    ' �J�����g��Workbook��temp�Ƃ������O�̃��[�N�V�[�g����������������폜����
    Dim wsName As String: wsName = "temp"
    Dim r As Boolean
    If KzIsWorksheetPresentInWorkbook(ThisWorkbook, wsName) Then
        r = KzDeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    End If
    'Act:
    ' temp���[�N�V�[�g��ǉ�����
    r = KzCreateWorksheetInWorkbook(ThisWorkbook, wsName)
    'Assert
    ' temp���[�N�V�[�g���ł��Ă��邱�Ƃ��m�F����
    Assert.IsTrue KzIsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
End Sub

'@TestMethod("KzDeleteWorksheetInWorkbook���e�X�g����")
Private Sub Test_KzDeleteWorksheetInWorkbook()
    'Arrange
    ' �J�����g��Workbook�Ɉꎞ�I�ȃ��[�N�V�[�g��}������A
    ' �V�[�g�̖��O�� temp �Ƃ���
    Dim wsName As String: wsName = "temp"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' �}���������[�N�V�[�g���폜����
    Dim r As Boolean
    r = KzDeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    ' �ꎞ�I�ɑ}���������[�N�V�[�g�����͂⑶�݂��Ȃ����Ƃ��m�F����A
    Assert.IsFalse KzIsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
End Sub



'@TestMethod("KzFetchWorksheetFromWorkbook�����j�b�g�e�X�g����")
Private Sub Test_KzFetchWorksheetFromWorkbook()
    On Error GoTo TestFail
    'Arrange
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "�e�X�g�f�[�^"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "copy"
    'Act
    Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsTrue KzIsWorksheetPresentInWorkbook(targetWorkbook, "copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("KzFetchWorksheetFromWorkbook��Err�𓊂���ꍇ")
Private Sub Test_KzFetchWorksheetFromWorkbook_shouldThrowErr()
    On Error GoTo TestFail
    KzUtil.KzCls
    'Arrange
    '���͂Əo�͂ɓ������[�N�u�b�N�̓����V�[�g���w�肵����G���[�ɂȂ��
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "�e�X�g�f�[�^"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "�e�X�g�f�[�^"
    'Act
    Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsFalse KzIsWorksheetPresentInWorkbook(targetWorkbook, "copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    'Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    'Err��raise�����͂��Ƃ킩���Ă���̂ŁA�������Ȃ��ŁA�Â��ɏI������ׂ�
End Sub

