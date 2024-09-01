Attribute VB_Name = "���[���󋵃`�F�b�N"
Option Explicit

'�{Sub�́A�X����Ȉ��̉�����{�N�x�̉������łɔ[�߂����ǂ�����
'�`�F�b�N���邱�Ƃ�ړI�Ƃ���v���O�����ł���B
'Visual Basic for Application����ɂ���Ď�������Ă���B

'�{Sub�́A�X����Ȉ��̉������Excel�t�@�C���ƌ����o�[��Excel�t�@�C���Ƃ�
'�ӂ���Excel�t�@�C���̓��e���Ƃ炵���킹�邱�ƂŕK�v�ȏ��𓱂��o���B

'�{���[�N�u�b�N�Ɂu�O���t�@�C���̃p�X�v���[�N�V�[�g��������
'���̂Ȃ��ɂӂ���Excel�t�@�C���̃p�X�������Ă���B

'�{Sub�͉������̃��[�N�u�b�N����{�N�x�̉���ꗗ�̃��[�N�V�[�g���R�s�[����
'�{���[�N�u�b�N�̂Ȃ��Ɂuwork�������v�Ƃ������O�̃��[�N�V�[�g�����B
'�{Sub�͊O���ɂ��郏�[�N�u�b�N��ǂނ����ɂ��āA���������Ȃ��B
'�{Sub�͊e���������[�t�������ǂ����̏����uwork�������v���[�N�V�[�g�ɏ������ށB

'�{Sub�͌����o�[����Excel��READ ONLY�ŎQ�Ƃ���B�����o�[���ɂ͂��������������݂��Ȃ��B

'�{���W���[�������s�������ʂƂ��Ă�����M����Ɂ~�̈󂪕t����ꂽ�Ƃ���B
'���ꂾ����M��������[�ł���Ɣ��f����̂͊낤���B
'�������ԈႦ�Ă����ώ��炾���璍�ӂ���B
'�l�Ԃ������o�[���̃f�[�^���悭�ǂ�Ŋm���߂悤�B
'�����o�[�����Ԉ���Ă��邩������Ȃ��B�������ɂ��肤��B
'�����o�[���̊ԈႢ�̂����ŁA����M���������U�荞��ł����Ƃ�������
'�{���W���[����ǂݎ�邱�Ƃ��ł��Ȃ�����������������Ȃ��B
'����ł���搶���ɖ��f�������Ȃ��悤�A�\���ɒ��ӂ���B

Public Sub Main()

    Dim modName As String: modName = "���[���󋵃`�F�b�N"
    Dim procName As String: procName = "Main"

    '�C�~�f�B�G�C�g�E�E�C���h�E������
    Call BbLog.Clear
    
    
    Call BbLog.Info(modName, procName, "���[���󋵃`�F�b�N���s���܂�")
    
    '�������Excel�t�@�C���̃p�X
    Dim memberFile As String: memberFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B2")
    Call BbLog.Info(modName, procName, "�������: " & memberFile)
    
    '�����o�[��Excel�t�@�C���̃p�X
    Dim cashbookFile As String: cashbookFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B3")
    Call BbLog.Info(modName, procName, "�����o�[��: " & cashbookFile)
    
    
    '============================================================================
    '�O���ɂ���������Excel�t�@�C���� "R6�N�x" �V�[�g���J�����g�̃��[�N�u�b�N��
    '�R�s�[����B"work�������"�V�[�g�������B���̓��e��ListObject�Ƃ��Ď��o���B
    Dim memberTable As ListObject
    Set memberTable = MbMemberTableUtil.FetchMemberTable(memberFile, "R6�N�x", ThisWorkbook)
    Call BbLog.Info(modName, procName, "����̐l�� memberTable.ListRows.Count=" & memberTable.ListRows.Count)
    
    'ListObject�̉E�[�ɗ��ǉ�����B��̖��O���u���[���󋵁v�Ƃ���
    
    
    'work�����o�[�����[�N�V�[�g�������Cashbook�I�u�W�F�N�g��͂�
    Dim cb As Cashbook
    Set cb = OpenCashbook()
    Call BbLog.Info(modName, procName, "�����o�[���̍s�� cb.Count=" & cb.Count)
    
    '�`�F�b�N�̑ΏۂƂ��ׂ��J�n���ƏI�������w�肵��������CashSelector�I�u�W�F�N�g���擾����
    Dim cs As CashSelector: Set cs = CbFactories.CreateCashSelector(cb, #4/1/2024#, #3/31/2025#)
    
    '�e���������[���������ǂ������ׂ�work�������ɏ�������
    
    '�������̑S�s�ɂ��ă��[�v
    Dim i As Long
    For i = 1 To memberTable.ListRows.Count
        '���������Ǝ����J�i�̂Q�Z���Ɏ��������Ă���s�܂薼��Ƃ��ėL���ȍs��I��
        Dim nameKanji As Variant: Set nameKanji = memberTable.ListColumns("����").DataBodyRange(i)
        Dim nameKana As Variant: Set nameKana = memberTable.ListColumns("�����J�i").DataBodyRange(i)
        Dim entitlement As Variant: Set entitlement = memberTable.ListColumns("���i").DataBodyRange(i)
        If Not nameKanji.Value = "" And Not nameKana = "" Then
            '���̐l�̎����J�i���L�[�Ƃ��Č����o�[������������
            Dim csList As CashList: Set csList = FindPaymentBy(cs, nameKana)
            '���i��A�̐l�AB�̐l�AC�̐l�AD�̐l�ɂ��Č����o�[���ƏƂ炵���킹��
            If entitlement = "A" Or entitlement = "B" Or entitlement = "C" Or entitlement = "D" Then
                If csList.Count = 1 Then
                    '�ʏ�ǂ���
                    Dim payedAt As String: payedAt = "R" & csList.Items(1).YY & "/" & csList.Items(1).MM & "/" & csList.Items(1).DD
                    Call PrintFinding(i, nameKana, entitlement, "�� " & payedAt)
                    Call RecordFindingIntoMemberTable(memberTable, i, "�� " & payedAt)
                ElseIf csList.Count > 1 Then
                    '���ꖼ�`�l����2���ȏ�̓�������B���������B�H���o�͂���B
                    Call PrintFinding(i, nameKana, entitlement, csList.Count & "?")
                    Call RecordFindingIntoMemberTable(memberTable, i, csList.Count & "?")
                    
                    '�ЂƂ�̉�����l�Ƃ��ĐU�荞�񂾂ق��ɋΖ����@����@���`�œ��Y����̉���U�荞�񂾃P�[�X���������B
                    '�ʁX�̖��`�l����̐U��������A���̃v���O�������d�������o���邱�Ƃ͂ł��Ȃ������B
                    '�o�[������S������҂͖��[�����o���邾���łȂ��A�ڂ��Â炵�ďd���[�������o����K�v������B
                    
                Else
                    '�[���Ȃ�����[�̉\������B�~���o�͂���B
                    Call PrintFinding(i, nameKana, entitlement, "�~")
                    Call RecordFindingIntoMemberTable(memberTable, i, "�~")
                End If
            ElseIf entitlement Like "*�O��*" Then
                '���i��B�O��AC�O��ɂ��Ă͌����o�[���Ƃ̏ƍ����X�L�b�v����
                '�O���ǎ����ǂ��S���Ԃ�܂Ƃ߂ĐU�荞�ނ���l�P�ʂ̏ƍ����ł��Ȃ��B�܂�������A�Ƃ����킯��
                Call PrintFinding(i, nameKana, entitlement, "�Z")
                Call RecordFindingIntoMemberTable(memberTable, i, "�Z")
            Else
                '���i���Ə��A�މ�A���̑��̏ꍇ�͌����o�[�����`�F�b�N����������o�͂���
                Call PrintFinding(i, nameKana, entitlement, "��")
                Call RecordFindingIntoMemberTable(memberTable, i, "��")
            End If
        End If
    Next i
    Call BbLog.Info(modName, procName, "���[���󋵃`�F�b�N���������܂���")
End Sub

'�����o�[�����[�N�V�[�g�ɊO������f�[�^�����[�h����Cashbok�I�u�W�F�N�g��Ԃ�
Private Function OpenCashbook() As Cashbook
    Dim wb As Workbook
    Set wb = Workbooks.Open(BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B3"))
    Dim sheetName As String: sheetName = "�����o�[��"
    Dim tableId As String: tableId = "CashbookTable1"
    Dim cb As Cashbook: Set cb = CbFactories.CreateCashbook(wb, sheetName, tableId)
    Set OpenCashbook = cb
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
End Function


Private Function FindPaymentBy(ByVal cs As CashSelector, ByVal nameKana As String)
    Dim csA, csB, csC, csD As CashList
    Set csA = cs.SelectCashListByMatchingDescription(AccountType.Income, "���", "A���", nameKana)
    Set csB = cs.SelectCashListByMatchingDescription(AccountType.Income, "���", "B���", nameKana)
    Set csC = cs.SelectCashListByMatchingDescription(AccountType.Income, "���", "C���", nameKana)
    Set csD = cs.SelectCashListByMatchingDescription(AccountType.Income, "���", "D���", nameKana)
    If csA.Count > 0 Then
        Set FindPaymentBy = csA
    ElseIf csB.Count > 0 Then
        Set FindPaymentBy = csB
    ElseIf csC.Count > 0 Then
        Set FindPaymentBy = csC
    ElseIf csD.Count > 0 Then
        Set FindPaymentBy = csD
    Else
        Set FindPaymentBy = CbFactories.CreateEmptyCashList()
    End If
End Function


Private Sub PrintFinding(ByVal i As Long, _
                            ByVal nameKana As String, _
                            ByVal entitlement As String, _
                            ByVal status As String)
    Dim message As String
    message = "|" & i & "|" & nameKana & "|" & entitlement & "|" & status & "|"
    Call BbLog.Info("���[���󋵃`�F�b�N", "Main", message)
End Sub
                    


Private Sub RecordFindingIntoMemberTable(ByVal memberTable As ListObject, _
                                    ByVal i As Long, _
                                    ByVal status As String)
    '�������e�[�u����i�Ԗڂ̍s�ɔ��茋�ʂ��������ށB
    memberTable.ListColumns("���[����").DataBodyRange(i).Value = status
    'status�����[�Ȃ�΂��̍s�̕�����ԐF�ɕύX����
    If status Like "�~" Then
        With memberTable.ListColumns("����").DataBodyRange(i).Font
            .Color = RGB(255, 64, 64)
            .Underline = True
        End With
    End If
End Sub

