Attribute VB_Name = "IDPW�ʒm����Word�����𐶐�"
Option Explicit

' �{Sub�́A�X����Ȉ��̉����l��l���ɗX������Word�����𐶐����܂��B
' �u���Ȃ����g�o�̉���y�[�W�ɃT�C���C������̂ɕK�v�Ȃh�c�ƃp�X���[�h�͂���ł��v�Ƃ��������B
' �{Sub�͊O���ɂ���X����Ȉ���������Excel�t�@�C����ǂݍ��ށB
' ����ɂ͊e����̎����Ƃh�c�ƃp�X���[�h�������Ă���B
' �e���v���[�g�Ƃ��Ă�Word�t�@�C����ǂݍ��݁A�v���[�X�z���_�[�Ƃ��Ă̋L�q�i${����} �Ȃǁj��
' Excel����E�����f�[�^�Œu������B���������l�����J��Ԃ��āA�l������Word�t�@�C�����o�͂���B

Public Sub IDPW�ʒm�����𐶐�()

    ' �C�~�f�B�G�C�g�E�E�C���h�E������
    Call BbLog.Clear
    
    
    Debug.Print ("ID/PW��ʒm����Word�����𐶐����܂�")
    
    ' �������Excel�t�@�C���̃p�X
    Dim memberFile As String: memberFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B2")
    Debug.Print ("�������: " & memberFile)
    
    ' �e���v���[�g�Ƃ��Ă�Word�t�@�C���̃p�X
    Dim templateFile As String: templateFile = _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B3")
    Debug.Print ("�e���v���[�g: " & templateFile)
    
    ' �o�̓t�H���_�̃p�X
    Dim outDir As String: outDir = BbFile.AbsolutifyPath( _
        ThisWorkbook.Path, _
        BbUtil.ResolveExternalFilePath(ThisWorkbook, "�O���t�@�C���̃p�X", "B4"))
    Debug.Print ("�o�̓t�H���_: " & outDir)
    
    ' �o�͐�t�H���_���������܂����݂��Ă��Ȃ���������
    Call BbFile.EnsureFolders(outDir)

    ' BbDocTransformer�C���X�^���X�𐶐�����
    Dim DT As BbDocTransformer: Set DT = BbDocTransformerFactory.CreateDocTransformer()
    ' Word�A�v���P�[�V�����̃C���X�^���X��^����
    Dim WordApp As Word.Application: Set WordApp = CreateObject("Word.application")
    ' DocTransformer������������
    Call DT.Initialize(WordApp)

    ' �O���ɂ���������Excel�t�@�C������V�[�g���R�s�[���Ď�荞�݁A���̒��ɂ����������ListObject�Ƃ��ĂƂ肾��
    Dim memberTable As ListObject
    Set memberTable = MbMemberTableUtil.FetchMemberTable(memberFile, "R6�N�x", ThisWorkbook)
    Debug.Print "memberTable.ListRows.Count=" & memberTable.ListRows.count
    
    
    ' �������̍s����������
    Dim max As Long: max = 300  ' �e�X�g����Ƃ��ɂ�max������������(3�Ƃ�)�ɂ��A�{�Ԃɂ͑���������傫�������ɂ���
    Dim count As Long: count = 0
    
    Dim i As Long
    For i = 1 To memberTable.ListRows.count
        If i < max Then
            ' ����̎�����ID��PW�̃f�[�^�����o��
            Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "����", Trim(memberTable.ListColumns("����").DataBodyRange(i))
            dict.Add "�����J�i", Trim(memberTable.ListColumns("�����J�i").DataBodyRange(i))
            dict.Add "ID", Trim(memberTable.ListColumns("HP��ID").DataBodyRange(i))
            dict.Add "PW", Trim(memberTable.ListColumns("HP�̃p�X���[�h").DataBodyRange(i))
            
            '���������Ǝ����J�i�̂Q�Z���Ɏ��������Ă���s�܂薼��Ƃ��ėL���ȍs��I��
            If Not dict("����") = "" And Not dict("�����J�i") = "" Then
                
                Debug.Print (dict("����") & " " & dict("�����J�i") & " " & dict("ID") & " " & dict("PW"))
                
                ' �V�������Word�t�@�C���̖��O�����߂�
                Dim r As String: r = outDir & "\" & "IDPW_" & dict("�����J�i") & ".docx"
                Debug.Print r
                
                ' Word�h�L�������g��ϊ����鏈�������s����
                Call DT.Transform(templateFile, dict, r)
            
            End If
        End If
    Next i
    
    ' Word�A�v���P�[�V���������
    WordApp.Quit
    Set WordApp = Nothing
    
    Debug.Print "�I�����܂����B"
    MsgBox "�o�͐�: " & outDir
    
End Sub

