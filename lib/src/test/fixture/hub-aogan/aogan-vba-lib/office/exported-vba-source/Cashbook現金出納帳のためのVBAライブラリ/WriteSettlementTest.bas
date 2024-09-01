Attribute VB_Name = "WriteSettlementTest"
Option Explicit
Option Private Module

' TestWriteSettlement : AccoutSum���W���[�����e�X�g����

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

'==============================================================================

'@TestMethod("�y�����o�[�L�^�z���[�N�V�[�g�����BGet���v�֐����e�X�g����B")
Private Sub Test�����o�[�L�^�����()
    On Error GoTo TestFail
    '�ߘa5�N�x�\�Z�i�āj�E�ߘa4�N�x���Z�i�āj.xlsm���[�N�u�b�N�ɍ�肱�܂ꂽ
    '���Z���쐬���W���[���̃R�[�h�Ɠ������̂�Rubberduck�̐��䉺�Ŏ��s����B

    Call BbLog.Clear      ' �C�~�f�B�G�C�g�E�B���h�E������������
    Call BbLog.Info("TestWriteSettlement", "Test�����o�[�L�^�����", "�����o�[�L�^���[�N�V�[�g���X�V���܂�")
    
    '���̃��[�N�u�b�N�̂Ȃ��Ɂy�����o�[�L�^�z���[�N�V�[�g�����
    Call WriteSettlement.�����o�[�L�^���[�N�V�[�g��������΍��(ThisWorkbook, "�����o�[�L�^")
    Dim ws�����o�[�L�^ As Worksheet
    Set ws�����o�[�L�^ = ThisWorkbook.Worksheets("�����o�[�L�^")
    
    '�����o�[�����[�N�u�b�N���J��
    Dim sourcePath As String
    sourcePath = BbUtil.ResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2")
    Call BbLog.Info("TestWriteSettlement", "Test�����o�[�L�^�����", "����: " & sourcePath)
    
    Dim wb�����o�[�� As Workbook: Set wb�����o�[�� = Workbooks.Open(sourcePath)
    Dim wsSource As Worksheet: Set wsSource = wb�����o�[��.Worksheets("�����o�[��")
    
    '�����o�[���̃f�[�^��I�ʂ��Ȃ���import����[�e�[�u�������o�[�L�^]�ɓ]�L����
    '�ߘa�S�N�S���P������ߘa�T�N�R���R�P���܂ł̓��o����I������
    '���x�񍐒P�ʂ��u���k�u���b�N�u�K��v�ł��郌�R�[�h�����O����B
    Call WriteSettlement.���o���L�^����荞��(wsSource, ws�����o�[�L�^, _
                                periodStart:=#4/1/2022#, periodEnd:=#3/31/2023#, _
                                ofReportingUnit:="���k�u���b�N�u�K��", positiveLike:=False)
    
    '[�����o�[�L�^]�e�[�u���̃f�[�^�s����בւ���A����Ȗږ��̖��ׂ�ǂݎ���悤��
    Call WriteSettlement.���o���L�^���\�[�g����(ws�����o�[�L�^)
    
    '[�e�[�u������Ȗڂ��Ƃ̏��v]���X�V����
    Call WriteSettlement.���v�̕\�����(ws�����o�[�L�^)
    
    ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肵��
    Application.DisplayAlerts = False
    '�����o�[�����[�N�u�b�N�����
    wb�����o�[��.Close
    Set wb�����o�[�� = Nothing
    
    Call BbLog.Info("TestWriteSettlement", "Test�����o�[�L�^�����", "�����o�[�L�^���[�N�V�[�g���X�V���܂���")


    'Get���v�֐����e�X�g����
    Dim val As Long
    val = WriteSettlement.Get���v("�x�o/�c����/�c����")
    Assert.IsTrue val > 0, "Get���v()���[����Ԃ���"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
