Attribute VB_Name = "�N���[���󋵃`�F�b�N"
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

    '�C�~�f�B�G�C�g�E�E�C���h�E������
    Call KzCls
    
    Debug.Print ("���[���󋵃`�F�b�N���s���܂�")
    
    '�������Excel�t�@�C���̃p�X
    Dim memberFile As String: memberFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B2")
    Debug.Print ("�������: " & memberFile)
    
    '�����o�[��Excel�t�@�C���̃p�X
    Dim cashbookFile As String: cashbookFile = _
        KzUtil.KzResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B3")
    Debug.Print ("�����o�[��: " & cashbookFile)
    
    
    '============================================================================
    '�O���ɂ���������Excel�t�@�C���� "R6�N�x" �V�[�g���J�����g�̃��[�N�u�b�N��
    '�R�s�[����B"work�������"�V�[�g�������B���̓��e��ListObject�Ƃ��Ď��o���B
    Dim memberTable As ListObject
    Set memberTable = AoMemberUtils.FetchMemberTable(memberFile, "R6�N�x", ThisWorkbook)
    Set memberTable = OpenMemberTable("R6�N�x", True)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.Count
    
    'work�����o�[�����[�N�V�[�g�������Cashbook�I�u�W�F�N�g��͂�
    Dim cb As Cashbook
    Set cb = OpenCashbook()
    Debug.Print "cb.Count=" & cb.Count
    Dim cs As CashSelector: Set cs = CreateCashSelector(cb, #4/1/2023#, #3/31/2024#)
    
    '�e���������[���������ǂ������ׂ�work�������ɏ�������
    Dim i As Long
    '�������̑S�s�ɂ��ă��[�v
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
    
End Sub

Private Function OpenMemberTable(Optional ByVal sheetName As String = "R6�N�x", Optional ByVal renew As Boolean = False) As ListObject
    '�}�X�^�̉�����냏�[�N�u�b�N���J���Ė���̃��[�N�V�[�g���J�����g�̃��[�N�u�b�N�Ɏ�荞��
    Dim targetWorkbook As Workbook: Set targetWorkbook = ThisWorkbook
    Dim targetSheetName As String: targetSheetName = "work�������"
    'work������냏�[�N�V�[�g���܂����݂��Ȃ��A�܂���
    '���[�N�V�[�g�����łɑ��݂��邪renew�p�����[�^��True�Ȃ��
    If (Not KzVerifyWorksheetExists(targetSheetName)) Or _
        (KzVerifyWorksheetExists(targetSheetName) And renew = True) Then
        Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(mbGetPathOfAoganMembers())
        Dim sourceSheetName As String: sourceSheetName = sheetName
        Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
        '������냏�[�N�u�b�N�����
        Application.DisplayAlerts = False   '�u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ�
        sourceWorkbook.Close
        Application.DisplayAlerts = True
    End If
    '�J�����g�̃��[�N�u�b�N�Ɏ�荞�񂾃��[�N�V�[�g�̂Ȃ��̃e�[�u����͂�
    Dim ws As Worksheet: Set ws = targetWorkbook.Worksheets(targetSheetName)
    Dim tbl As ListObject: Set tbl = ws.ListObjects("MembersTable13")
    
    '�������̃e�[�u���̉E�[�񂪁u��4�v�Ƃ������O�ł���Ƃ�����u���[���󋵁v�ɕύX����
    tbl.ListColumns(tbl.ListColumns.Count).Name = "���[����"
    
    '�񕝂𒲐�
    tbl.ListColumns("�����J�i").Range.EntireColumn.AutoFit
    tbl.ListColumns("�Ζ��於").Range.EntireColumn.AutoFit
    tbl.ListColumns("�ٓ�").Range.EntireColumn.AutoFit
    tbl.ListColumns("���[����").Range.EntireColumn.AutoFit
    
    '�d�v�łȂ�����\���ɂ���B�V�[�g����������Ƃ��ɕ֗��Ȃ悤��
    tbl.ListColumns("�N��").Range.EntireColumn.Hidden = True
    tbl.ListColumns("��oNo").Range.EntireColumn.Hidden = True
    tbl.ListColumns("��o�^��").Range.EntireColumn.Hidden = True
    tbl.ListColumns("�����o�^").Range.EntireColumn.Hidden = True
    tbl.ListColumns("�Ζ���").Range.EntireColumn.Hidden = True
    tbl.ListColumns("�Ζ���Z��").Range.EntireColumn.Hidden = True
    tbl.ListColumns("�ΐ�TELNo").Range.EntireColumn.Hidden = True
    tbl.ListColumns("���").Range.EntireColumn.Hidden = True
    tbl.ListColumns("����Z��").Range.EntireColumn.Hidden = True
    tbl.ListColumns("����TELNo").Range.EntireColumn.Hidden = True
    tbl.ListColumns("�g�єԍ�").Range.EntireColumn.Hidden = True
    
    Set OpenMemberTable = tbl
End Function

Private Function OpenCashbook() As Cashbook
    Dim wb As Workbook: Set wb = Workbooks.Open(KzResolveExternalFilePath(ThisWorkbook, "�����o�[���t�@�C���̃p�X", "B2"))
    Dim sheetName As String: sheetName = "�����o�[��"
    Dim tableId As String: tableId = "CashbookTable1"
    Dim cb As Cashbook: Set cb = CreateCashbook(wb, sheetName, tableId)
    Set OpenCashbook = cb
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
End Function

'Private Function OpenCashbook(Optional ByVal renew As Boolean = False) As Cashbook
'    '�}�X�^�[�̌����o�[�����[�N�u�b�N���Ђ炢�Č����o�[���̃��[�N�V�[�g���J�����g�̃��[�N�u�b�N�Ɏ�荞��
'    Dim targetWorkbook As Workbook: Set targetWorkbook = ThisWorkbook
'    Dim targetSheetName As String: targetSheetName = "work�����o�[��"
'    'work������냏�[�N�V�[�g���܂����݂��Ȃ��A�܂���
'    '���[�N�V�[�g�����łɑ��݂��邪renew�p�����[�^��True�Ȃ��
'    If (Not KzVerifyWorksheetExists(targetSheetName)) Or _
'        (KzVerifyWorksheetExists(targetSheetName) And renew = True) Then
'        Dim sourceWorkbook As Workbook: Set sourceWorkbook = Workbooks.Open(GetPathOfAoganCashbook)
'        Dim sourceSheetName As String: sourceSheetName = "�����o�[��"
'        Call KzFetchWorksheetFromWorkbook(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName)
'        '�����o�[�����[�N�u�b�N�����
'        Application.DisplayAlerts = False   '�u�ύX���e��ۑ����܂����v�_�C�A���O��\�����Ȃ�
'        sourceWorkbook.Close
'        Application.DisplayAlerts = True
'    End If
'    '�J�����g�̃��[�N�u�b�N�Ɏ�荞�񂾃��[�N�V�[�g�̂Ȃ��̃e�[�u����͂�
'    Dim cb As Cashbook: Set cb = CreateCashbook(targetWorkbook, targetSheetName)
'    '
'    Set OpenCashbook = cb
'End Function



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
        Set FindPaymentBy = CreateEmptyCashList()
    End If
End Function


Private Sub PrintFinding(ByVal i As Long, _
                            ByVal nameKana As String, _
                            ByVal entitlement As String, _
                            ByVal status As String)
    Debug.Print "|"; i & "|" & nameKana & "|" & entitlement & "|" & status & "|"
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

