Attribute VB_Name = "���[�҂ւ̓�����쐬"
Option Explicit

'������쐬����
'work������냏�[�N�V�[�g�̂Ȃ���Table�̉��[���󋵂��~�ƈ�Â����Ă�������I�����A������쐬����B
'MakeLetter�v���V�W���͊O���t�@�C����������������[�h���邪�A����ƈقȂ�AMakeReminder�N���X�͊O���t�@�C�������[�h���Ȃ��B
'MakeReminder�v���V�W���́u���[���󋵃`�F�b�N�v���W���[�����쐬���X�V�����uwork�������v���[�N�V�[�g�v���Q�Ƃ���B
'MakeReminder�v���V�W���́uwork�������v���[�N�V�[�g�́u���[���󋵁v��𒲂ׁA�����Ɂ���~�Ȃǂ̗L���ȕ�����
'�L������Ă��邱�Ƃ��`�F�b�N����B�������L���ȕ������L������Ă��Ȃ���΃G���[���b�Z�[�W��\�����ďI������B


Public Sub MakeReminder()
    
    Call BbLog.Clear
    Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", "�n�߂܂�")
    
    '����̃e���v���[�g�Ƃ��Ă�Word�t�@�C���̃p�X
    Dim templateFile As String: templateFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B6")
    Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", "�e���v���[�g: " & templateFile)
    
    '�o�͐�t�H���_�̃p�X
    Dim outDir As String: outDir = BbFile.AbsolutifyPath( _
        ThisWorkbook.Path, _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B7"))
    Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", "�o�͐�t�H���_: " & outDir)
    
    ' �o�͐�t�H���_�����łɂ�������폜����
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(outDir) Then
        FSO.DeleteFolder outDir
    End If
    ' �o�͐�t�H���_�����������Ǎ��
    Call BbFile.EnsureFolders(outDir)
    
    ' BbDocTransformer�C���X�^���X����������
    Dim DT As BbDocTransformer: Set DT = BbDocTransformerFactory.CreateDocTransformer()
    ' Word�A�v���P�[�V�����̃C���X�^���X��^����
    Dim WordApp As Word.Application: Set WordApp = CreateObject("Word.application")
    ' DocTrasnsformer������������
    Call DT.Initialize(WordApp)
    
    '=================================================================================
    ' ���̃��[�N�u�b�N�Ɂuwork�������v�V�[�g������͂��B���̒��g��ListObject�Ƃ��Ď��o��
    '
    Dim memberTable As ListObject
    Set memberTable = ThisWorkbook.Worksheets("work�������").ListObjects("MembersTable13")
    
    Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", _
                            "memberTable.ListRows.Count=" & memberTable.ListRows.count)
    
    ' ================================================================================
    ' �������̊e�s����������
    Dim max As Long: max = 300     '�e�X�g���ɂ͏���������(3�Ƃ�)�ɂ��đ����I��������
                                 '�{�Ԃɂ͑���������傫������(300�Ƃ�)�ɂ���
    Dim count As Long: count = 0

    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i < max Then
            ' ����̍s���玁���A�����J�i�A���i�A���[���󋵂����o��
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "����", Trim(memberTable.ListColumns("����").DataBodyRange(i))
            dict.Add "�����J�i", Trim(memberTable.ListColumns("�����J�i").DataBodyRange(i))
            dict.Add "���i", Trim(memberTable.ListColumns("���i").DataBodyRange(i))
            dict.Add "���[����", Trim(memberTable.ListColumns("���[����").DataBodyRange(i))
            Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", dict("�����J�i") & dict("���[����"))
    
            
            ' ���[���󋵂����L���Ȃ�΁i���[���󋵃`�F�b�N���܂����s����Ă��Ȃ����Ƃ��Ӗ�����̂Łj
            ' Err�𓊂��ďI������
            If dict("���[����") = "" Then
                Err.Raise 5000, "���[�҂ւ̓�����쐬.MakeReminder", _
                        dict("�����J�i") & "�̉��[���󋵂����L���B" & _
                        "���O�ɉ��[���󋵃`�F�b�N�����s����K�v������܂�"
            End If
            
            ' ���[���󋵂��~�ł�������ΏۂƂ��ē����Word�t�@�C�����쐬����B
            ' ���[���󋵂��~�ł͂Ȃ�����ɂ��Ă͐������Ȃ��B
            If dict("���[����") Like "�~" Then
                Dim msg As String: Let msg = dict("�����J�i") & " " & dict("����") & " " & dict("���i")
                Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", msg)

                ' �o��Word�t�@�C���̃p�X�����肵��
                Dim r As String: r = outDir & "\" & dict("�����J�i") & "_" & dict("����") & "_" & dict("���i") & ".docx"
                Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", "�o�̓t�@�C�� " & r)
            
                ' Word�h�L�������g��ϊ����鏈�������s����
                Call DT.Transform(templateFile, dict, r)
    
            End If
        End If
    Next i
    
    Call BbLog.Info("���[�҂ւ̓�����쐬", "MakeReminder", "�I���܂���")
    
End Sub
