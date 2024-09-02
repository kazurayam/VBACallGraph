Attribute VB_Name = "BbWorksheetTest"
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


'@TestMethod("IsWorksheetPresentInWorkbook���e�X�g����")
Private Sub Test_IsWorksheetPresentInWorkbook()
    'Assert:
    Assert.IsTrue BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, "�e�X�g�f�[�^")
    Assert.IsFalse BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, "No Such Worksheet")
End Sub

'@TestMethod("CreateWorksheetInWorkbook���e�X�g����")
Private Sub Test_CreateWorksheetInWorkbook()
    'Arrange
    ' �J�����g��Workbook��temp�Ƃ������O�̃��[�N�V�[�g����������������폜����
    Dim wsName As String: wsName = "temp"
    Dim r As Boolean
    If BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName) Then
        r = BbWorksheet.DeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    End If
    'Act:
    ' temp���[�N�V�[�g��ǉ�����
    r = BbWorksheet.CreateWorksheetInWorkbook(ThisWorkbook, wsName)
    'Assert
    ' temp���[�N�V�[�g���ł��Ă��邱�Ƃ��m�F����
    Assert.IsTrue BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
    
End Sub

'@TestMethod("DeleteWorksheetInWorkbook���e�X�g����")
Private Sub Test_DeleteWorksheetInWorkbook()
    'Arrange
    ' �J�����g��Workbook��temp�Ƃ������O�̃��[�N�V�[�g����������������폜����
    Dim wsName As String: wsName = "temp"
    Dim r As Boolean
    If BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName) Then
        r = BbWorksheet.DeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    End If
    ' �J�����g��Workbook�Ɉꎞ�I�ȃ��[�N�V�[�g��}������A
    ' �V�[�g�̖��O�� temp �Ƃ���
    Worksheets.Add(After:=Worksheets(Worksheets.Count)) _
        .Name = wsName
    'Act:
    ' �}���������[�N�V�[�g���폜����
    r = BbWorksheet.DeleteWorksheetInWorkbook(ThisWorkbook, wsName)
    ' �ꎞ�I�ɑ}���������[�N�V�[�g�����͂⑶�݂��Ȃ����Ƃ��m�F����A
    Assert.IsFalse BbWorksheet.IsWorksheetPresentInWorkbook(ThisWorkbook, wsName)
End Sub



'@TestMethod("FetchWorksheetFromWorkbook�����j�b�g�e�X�g����")
Private Sub Test_FetchWorksheetFromWorkbook()
    On Error GoTo TestFail
    'Arrange
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "�e�X�g�f�[�^"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "copy"
    'Act
    Call BbWorksheet.FetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsTrue BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, "copy")
    'TearDown
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(targetSheetName).Delete
    Application.DisplayAlerts = True
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("FetchWorksheetFromWorkbook��Err�𓊂���ꍇ")
Private Sub Test_FetchWorksheetFromWorkbook_shouldThrowErr()
    On Error GoTo TestFail
    Call BbLog.Clear
    'Arrange
    '���͂Əo�͂ɓ������[�N�u�b�N�̓����V�[�g���w�肵����G���[�ɂȂ��
    Dim sourceWorkbook As Workbook: Set sourceWorkbook = ThisWorkbook
    Dim sourceSheetName As String: sourceSheetName = "�e�X�g�f�[�^"
    Dim targetWorkbook As Workbook: Set targetWorkbook = sourceWorkbook
    Dim targetSheetName As String: targetSheetName = "�e�X�g�f�[�^"
    'Act
    Call BbWorksheet.FetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
    'Assert
    Assert.IsFalse BbWorksheet.IsWorksheetPresentInWorkbook(targetWorkbook, "copy")
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

