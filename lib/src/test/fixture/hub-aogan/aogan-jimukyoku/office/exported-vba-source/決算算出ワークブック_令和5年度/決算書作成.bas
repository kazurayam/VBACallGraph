Attribute VB_Name = "���Z���쐬"
Option Explicit

' ���Z���쐬 Module
'
'�X����Ȉ��̌����o�[�����[�N�u�b�N�ɋL�^���ꂽ���o���f�[�^����ȓ��͏��Ƃ���
'�w���Z���x�V�[�g�ɐ����𖄂߂邱�Ƃ��ŏI�ړI�Ƃ���B
'
'�w���s���x�V�[�g�̂Ȃ��ɊO���̃��[�N�u�b�N�̋�̓I�ȃp�X�������Ă���̂𗊂�Ƃ���
'�����o�[�����[�N�u�b�N��open����B
'
'�����o�[�����[�N�u�b�N�̂Ȃ����玟�̏����ɊY������s��I������B
'  1. �ߘa5�N4��1������ߘa6�N3��31���܂ł̊��Ԃɑ�������o���@����
'  2. �y���x�񍐒P�ʁz���u���k�u���b�N�u�K��v�ł͂Ȃ����o��
'�I�������s���w���o���L�^�x�V�[�g�̃e�[�u���ɃR�s�[����B
'�e�[�u���̍s������Ȗڂ̏����@���@�N�����̏����ŕ��בւ���B
'���בւ���Ɓw���o���L�^�x������Ȗڂ��Ƃ̖��׏��Ƃ��ēǂނ��Ƃ��ł���悤�ɂȂ�B

'�w���o���L�^�x�e�[�u���̃f�[�^���犨��Ȗږ��̏��v���Z�o��[����Ȗڂ��Ƃ̏��v]�e�[�u���ɏo�͂���B
'�w���Z���x���[�N�V�[�g�̃Z���ɎQ�Ǝ���ݒ肵�AFunction���Ăяo����B
'Function��[���v�e�[�u��]���������Ċ���Ȗڂ��Ƃ̏��v���z��Ԃ��B


'==============================================================================
Public Sub ���Z���쐬()

    Call KzUtil.KzCls      ' �C�~�f�B�G�C�g�E�B���h�E������������
    Call KzUtil.KzLog("���Z���쐬", "���Z���쐬", "�����o�[�L�^���[�N�V�[�g���X�V���܂�")
    
    '���̃��[�N�u�b�N�̂Ȃ��Ɂy�����o�[�L�^�z���[�N�V�[�g�����
    Call CashbookPrj.WriteSettlement.�����o�[�L�^���[�N�V�[�g��������΍��(ThisWorkbook, "�����o�[�L�^")
    
    Dim ws�����o�[�L�^ As Worksheet
    Set ws�����o�[�L�^ = ThisWorkbook.Worksheets("�����o�[�L�^")
    
    '�����o�[�����[�N�u�b�N���J��
    Dim sourcePath As String
    sourcePath = KzUtil.KzResolveExternalFilePath(ThisWorkbook, "���s��", "B2")
    Call KzUtil.KzLog("���Z���쐬", "���Z���쐬", "����: " & sourcePath)
    
    Dim wb�����o�[�� As Workbook: Set wb�����o�[�� = Workbooks.Open(sourcePath)
    Dim wsSource As Worksheet: Set wsSource = wb�����o�[��.Worksheets("�����o�[��")
    
    '�����o�[���̃f�[�^��I�ʂ��Ȃ���import����[�e�[�u�������o�[�L�^]�ɓ]�L����
    '�ߘa5�N�S���P������ߘa6�N�R���R�P���܂ł̓��o����I������
    '���x�񍐒P�ʂ��u���k�u���b�N�u�K��v�ł��郌�R�[�h�����O����B
    Call CashbookPrj.WriteSettlement.���o���L�^����荞��(wsSource, ws�����o�[�L�^, _
                                periodStart:=#4/1/2023#, periodEnd:=#3/31/2024#, _
                                ofReportingUnit:="���k�u���b�N�u�K��", positiveLike:=False)
    
    '[�����o�[�L�^]�e�[�u���̃f�[�^�s����בւ���A����Ȗږ��̖��ׂ�ǂݎ���悤��
    Call CashbookPrj.WriteSettlement.���o���L�^���\�[�g����(ws�����o�[�L�^)
    
    '[�e�[�u������Ȗڂ��Ƃ̏��v]���X�V����
    Call CashbookPrj.WriteSettlement.���v�̕\�����(ws�����o�[�L�^)
    
    ' �u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ��悤�ɐݒ肵��
    Application.DisplayAlerts = False
    '�����o�[�����[�N�u�b�N�����
    wb�����o�[��.Close
    Set wb�����o�[�� = Nothing
    
    Call KzUtil.KzLog("���Z���쐬", "���Z���쐬", "�����o�[�L�^���[�N�V�[�g���X�V���܂���")
    
End Sub

