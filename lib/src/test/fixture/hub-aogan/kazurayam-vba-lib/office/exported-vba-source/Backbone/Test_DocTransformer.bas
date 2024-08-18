Attribute VB_Name = "Test_DocTransformer"
Option Explicit


' DocTransformer�N���X�����s���ē�����m�F���܂��ARubberduck�ɂ�鎩�����e�X�g�ł͂Ȃ��蓮�Ŏ��s����
' ���͂Ƃ��Ă�KeyValuePair���R�^����A���̂R�̓R�[�h�̂Ȃ��ɌŒ�l�Ƃ��ď����Ă���

Public Sub Test()
    ' �C�~�f�B�G�C�g�E�E�C���h�E�������B
    ' ����̎��s��Debug.Print���o�͂��郁�b�Z�[�W�����₷�����邽�߁B
    Call KzCls
    
    ' ���͂ƂȂ�e���v���[�g�Ƃ��Ă�Word�t�@�C�����i�[���ꂽ�t�H���_�̃p�X
    Dim TemplateFolder As String: TemplateFolder = KzFile.KzToLocalFilePath(ThisWorkbook.path & "\" & "data")
    ' �e���v���[�g�Ƃ��Ă�Word�t�@�C���̖��O
    Dim TemplateName As String: TemplateName = "�e���v���[�g.docx"
    
    ' �o�͂����Word�t�@�C�����i�[�����Ƃ��Ẵt�H���_�̃p�X
    Dim ResultFolder As String
    ' OneDrive�ŗL�̌`���̃t�@�C���p�X�Ȃ�Ε��ʂ̃t�@�C���p�X�ɕϊ�����
    ResultFolder = KzFile.KzToLocalFilePath(ThisWorkbook.path) & "\" & "build\Test_DocTransformer"
    ' �o�͐�t�H���_���������܂����݂��Ă��Ȃ���������
    Call KzFile.KzEnsureFolders(ResultFolder)
    
    ' �v���[�X�z���_�[�̖��O�ƒl�̑g
    Dim dict As Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "����", "�D�c�M��"
    dict.Add "�Ζ���", "���ǐ�N���j�b�N"
    dict.Add "���i", "A"
    
    ' �u�����v�Ɋ�Â��ďo��Word�t�@�C���̖��O�����߂�
    Dim ResultName As String: ResultName = dict("����") & ".docx"
        
    ' ����Word�t�@�C���Əo��Word�t�@�C���̃p�X�����肵��
    Dim t As String: t = TemplateFolder & "\" & TemplateName
    Dim r As String: r = ResultFolder & "\" & ResultName
        
    Debug.Print "�e���v���[�g=" & t
    Debug.Print "�o�̓t�@�C��=" & r
    
    ' Word�h�L�������g��ϊ�����N���X������������
    ' DocTransformer�C���X�^���X�𐶐�����
    Dim DT As DocTransformer
    Set DT = DocTransformerUtil.Create
    
    
    ' Word�A�v���P�[�V�����̃C���X�^���X�𐶐�
    Dim WordApp As Word.Application: Set WordApp = CreateObject("Word.application")
    Call DT.Initialize(WordApp)
    ' ���̓t�@�C���ƒu���f�[�^�Əo�̓t�@�C�����w�肵��Word�h�L�������g��ϊ����o�͂���
    Call DT.Transform(t, dict, r)
    
    Debug.Print "�ϊ����������s���܂���"
    
    ' =========================================================================
    ' ��n��
    Set DT = Nothing
    Set dict = Nothing
    Set WordApp = Nothing

End Sub


