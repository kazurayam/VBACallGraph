Attribute VB_Name = "Environment"
Option Explicit

Function mbGetPathOfAoganCashbook() As String
    ' ���̃u�b�N�̂Ȃ��̎��s�����[�N�V�[�g�̒���
    ' �X����Ȉ��̌����o�[��Excel�u�b�N�̃p�X�������Ă���B
    ' ���̃p�X�͂���Excel�u�b�N�̃p�X�����Ƃ��鑊�΃p�X�ł���B
    ' ���̃p�X�̒l�����[�N�V�[�g����ǂݏo���āA��΃p�X�ɕϊ����āA
    ' String�Ƃ��ĕԂ��B
    mbGetPathOfAoganCashbook = KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2")
End Function

Sub Test_mbGetPathOfAoganCashbook()
    ' GetPathOfAoganMembers���e�X�g����
    Call KzCls
    'Sub Cls1�� Cashbook�����o�[���̂��߂�VBA���C�u����.xlam �Ɋ܂܂�Ă���
    Debug.Print mbGetPathOfAoganCashbook()
End Sub

Function mbGetPathOfAoganMembers() As String
    'ThisWorkbook�̂Ȃ��́u�������t�@�C���̃p�X�v���[�N�V�[�g�̂Ȃ���
    '�X����Ȉ��̉������Excel�t�@�C���̃p�X�������Ă���B
    '���̃p�X��ThisWorkbook�̃p�X�����Ƃ��鑊�΃p�X�ł���B
    '���̒l��ǂ݂����Đ�΃p�X�ɕϊ�����String�Ƃ��ĕԂ��B
    mbGetPathOfAoganMembers = KzResolveExternalFilePath(ThisWorkbook, "�������t�@�C���̃p�X", "B2")
End Function

Sub Test_mbGetPathOfAoganMembers()
    ' GetPathOfAoganMembers���e�X�g����
    Call KzCls
    'Sub Cls1�� Cashbook�����o�[���̂��߂�VBA���C�u����.xlam �Ɋ܂܂�Ă���
    Debug.Print mbGetPathOfAoganMembers()
End Sub
