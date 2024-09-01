Attribute VB_Name = "��HP��IDPW�ꗗCSV�𐶐�"
Option Explicit


' ��HP�̉���̃y�[�W https://www.aomori-gankaikai.jp/member ��Basic�F�؂ɂ��A�N�Z�X�������{���Ă���B
' ���̍����ƂȂ�ID�ƃp�X���[�h�̑g�����^����CSV�t�@�C�����������Excel����o�͂���B

Public Sub CSV�𐶐�()

    ' �C�~�f�B�G�C�g�E�C���h�E������
    Call BbLog.Clear
    
    Debug.Print "��HP��ID�ƃp�X���[�h�̑g�̈ꗗ��CSV�t�@�C���ɏo�͂���"
    
    ' �������Excel�t�@�C���̃p�X
    Dim memberFile As String: memberFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B2")
    Debug.Print ("�������: " & memberFile)
    
    
    ' �o�͐�t�H���_�̃p�X
    Dim outDir As String: outDir = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B5")
    ' CSV�t�@�C���̏o�͐�p�X
    Dim CSV As String: CSV = outDir & "\web_account.csv"
    Debug.Print "�o�͐�: " + CSV
    
    ' �o�͐�t�H���_���������܂����݂��Ă��Ȃ���������
    Call BbFile.EnsureFolders(outDir)
    
    ' CSV�e�L�X�g���o�͂��邽�߂ɃX�g���[�����J��
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    Set ts = fs.CreateTextFile(CSV, True, False)
    
    ' Web�T�C�g�Ǘ��҂̔F�؏���1�s�ڂ�2�s�ڂɏ����B
    ' �Ȃ��Ȃ�Web�T�C�g�Ǘ��҂͊�Ȉ��̐�����ł͂Ȃ�����������Ɋ܂܂�Ă��Ȃ�����A����Ƃ��āB
    Dim plus1 As String: plus1 = "kazuaki001,a9ft5t72,""�Y�R�a���@������"""
    Call ts.Write(plus1)
    Call ts.Write(vbCrLf)
    
    Dim plus2 As String: plus2 = "aomoriken,gankaikai,""aomoriken / gankaikai"""
    Call ts.Write(plus2)
    Call ts.Write(vbCrLf)

    ' �O���ɂ���������Excel�t�@�C������V�[�g���R�s�[���Ď�荞�݁A
    ' ���̒��ɂ����������ListObject�Ƃ��ĂƂ肾��
    Dim memberTable As ListObject
    Set memberTable = MbMemberTableUtil.FetchMemberTable(memberFile, "R6�N�x", ThisWorkbook)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.count
    
    ' �������̍s����������
    Dim max As Long: max = 300   ' �e�X�g����Ƃ��ɂ�max�ɏ���������(3�Ƃ�)���Z�b�g���Ď��s���Ԃ�Z�k����
                                    ' �{�Ԃɂ͑���������傫������(300�Ƃ�)���Z�b�g����
    Dim count As Long: count = 0
    
    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i <= max Then
            ' ����̎�����ID��PW�̃f�[�^�����o��
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "����", Trim(memberTable.ListColumns("����").DataBodyRange(i))
            dict.Add "�����J�i", Trim(memberTable.ListColumns("�����J�i").DataBodyRange(i))
            dict.Add "HP��ID", Trim(memberTable.ListColumns("HP��ID").DataBodyRange(i))
            dict.Add "HP�̃p�X���[�h", Trim(memberTable.ListColumns("HP�̃p�X���[�h").DataBodyRange(i))
            '���������Ǝ����J�i�̂Q�Z���Ɏ��������Ă���s�܂薼��Ƃ��ėL���ȍs��I��
            If Not dict("����") = "" And Not dict("�����J�i") = "" Then
                Dim line As String: line = dict("HP��ID") & "," & dict("HP�̃p�X���[�h") & ",""" & dict("����") & """"
                Debug.Print line
                '�t�@�C����Write
                Call ts.Write(line)
                Call ts.Write(vbCrLf)
            End If
        End If
    Next i
    
    '�o�͐�t�@�C�����N���[�Y
    Call ts.Close
    
    ' ����������ʒm����A�o�͐��\������
    Call MsgBox("�o�͐�: " + CSV)
    
End Sub

