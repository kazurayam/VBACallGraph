Attribute VB_Name = "Relay"
Option Explicit

Private Const SHEETPASSWORD As String = "umeume0416"

'**
'* ����V�[�g�ɕی��K�p���邳���Ɏw�肷��p�X���[�h
'*
Public Property Get sheetPW()
    sheetPW = SHEETPASSWORD
End Property

'**
'* �u����v�V�[�g�́u����Ǘ��v�{�^�����N���b�N���ꂽ�炱��ShowUserForm���Ăяo���B
'* �u����e�X�g�p�v�V�[�g�́u����Ǘ��v�{�^��������������ShowUserForm���Ăяo���B
'*
Public Sub ShowUserForm()
    With UserForm1
        '�t�H�[����\������ʒu���w�肷��
        .StartUpPosition = 0
        .Left = 320
        .Top = 220
        
        '����Ǘ��t�H�[�����J��
        UserForm1.Show vbModeless
    
        '����V�[�g�̃f�[�^��UserForm1�̂Ȃ��Ƀ��[�h����
        UserForm1.LoadData ActiveSheet
    End With
End Sub
