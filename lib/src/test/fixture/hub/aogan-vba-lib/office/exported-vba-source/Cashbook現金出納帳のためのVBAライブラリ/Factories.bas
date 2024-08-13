Attribute VB_Name = "Factories"
Option Explicit

' Factories --- Cashbook�v���W�F�N�g�̃N���X���W���[���̂����O���v���W�F�N�g���C���X�^���X�𐶐����邽�߂ɗ��p����Function����

' �����Ő錾���ꂽFunction�����j�b�g�e�X�g����R�[�h��TestFactories���W���[���ɂ���

' ���̃v���W�F�N�g��Addin�܂�*.xlam�t�@�C���ɂ��āA�O���v���W�F�N�g�� �c�[�� > �Q�Ɛݒ� > �Q�� ��*.xlam�t�@�C����I�������Ƃ�
' �O���v���W�F�N�g�́@New Cashbook���邱�Ƃ��ł��Ȃ��B
' �Ƃ����̂�VBA�̃N���X��PublicNonCreatable�ȑ����������Ă���̂ŁA�O���v���W�F�N�g�̃R�[�h��New Cashbook���邱�Ƃ��ł��Ȃ��B
' ������*.xlam�̂Ȃ���Public�ȃt�@�N�g���֐����������Ă����B�O���v���W�F�N�g��Public�ȃt�@�N�g���֐��Ȃ�ΌĂяo�����Ƃ�
' �������B


Public Function CreateCashbook(ByVal wb As Workbook, _
                                ByVal sheetName As String, _
                                ByVal tableId As String) As Cashbook
    '�����Ƃ��ēn���ꂽWorkbook����͂Ƃ���Cashbook�I�u�W�F�N�g�𐶐����ĕԂ�
    '�O���v���W�F�N�g��Cashbook�I�u�W�F�N�g�𐶐����邽�߂�Public�Ȃ��̊֐����K�v��
    Debug.Print "wbFullName=" & wb.FullName
    Debug.Print "sheetName=" & sheetName
    Debug.Print "tableId=" & tableId
    
    Dim ws As Worksheet: Set ws = wb.Worksheets(sheetName)
    Dim tbl As ListObject: Set tbl = ws.ListObjects(tableId)
    Dim cb As Cashbook: Set cb = New Cashbook
    Call cb.Initialize(tbl)
    Set CreateCashbook = cb
End Function


Public Function CreateAccountsFinder(ByVal cb As Cashbook) As AccountsFinder
    '�����Ƃ��ēn���ꂽCashbook����͂Ƃ���AccountsFinder�I�u�W�F�N�g�𐶐����ĕԂ�
    '�O���v���W�F�N�g��AccountsFinder�I�u�W�F�N�g�𐶐����邽�߂�Public�Ȃ��̊֐����K�v��
    Dim accFinder As AccountsFinder: Set accFinder = New AccountsFinder
    Call accFinder.Initialize(cb)
    Set CreateAccountsFinder = accFinder
End Function



Public Function CreateCashSelector(ByVal cb As Cashbook) As CashSelector
    '�����Ƃ��ēn���ꂽCashbook����͂Ƃ���CashSelector�I�u�W�F�N�g�𐶐����ĕԂ�
    '�O���v���W�F�N�g��CashSelector�I�u�W�F�N�g�𐶐����邽�߂�Public�Ȃ��̊֐����K�v��
    Dim cs As CashSelector: Set cs = New CashSelector
    Call cs.Initialize(cb)
    Set CreateCashSelector = cs
End Function


Public Function CreateEmptyCashList() As CashList
    '�����[����CashList�I�u�W�F�N�g��Ԃ�
    'CashList�I�u�W�F�N�g��Private�Ɛ錾��������O�����W���[����New CashList()�ł��Ȃ��Ȃ����B
    '�����⊮���邽�߁B
    Dim cl As CashList
    Set cl = New CashList
    Set CreateEmptyCashList = cl
End Function
